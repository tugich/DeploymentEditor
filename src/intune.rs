//! Native Microsoft Intune uploader (Microsoft Graph, beta endpoint).
//!
//! Flow (same as the IntuneWin32App PowerShell module):
//! 1. sign in (app registration + client secret, or device code)
//! 2. POST   deviceAppManagement/mobileApps                         (or PATCH an existing app)
//! 3. POST   …/microsoft.graph.win32LobApp/contentVersions
//! 4. POST   …/contentVersions/{cv}/files  → wait for azureStorageUri
//! 5. PUT    Azure Storage blocks (6 MB) + block list, renewing the SAS URI when needed
//! 6. POST   …/files/{file}/commit with the encryption info → wait for commitFileSuccess
//! 7. PATCH  mobileApps/{id} committedContentVersion

use std::time::{Duration, Instant};

use anyhow::{anyhow, bail, Context, Result};
use serde_json::{json, Value};

use crate::intunewin::EncryptedPackage;

pub const DEFAULT_PUBLIC_CLIENT_ID: &str = "14d82eec-204b-4c2f-b7e8-296a70dab67e"; // Microsoft Graph Command Line Tools
pub const SCOPE_DELEGATED: &str = "https://graph.microsoft.com/DeviceManagementApps.ReadWrite.All offline_access";
pub const SCOPE_APP: &str = "https://graph.microsoft.com/.default";
const CHUNK_SIZE: usize = 6 * 1024 * 1024;
const KEYRING_SERVICE: &str = "DeploymentEditor-Intune";

// ---------------------------------------------------------------------------- endpoints & http

#[derive(Clone, Debug)]
pub struct Endpoints {
    pub login: String,
    pub graph: String,
}

impl Default for Endpoints {
    fn default() -> Self {
        // A mock server can be used in debug builds (tests); release builds always talk to Microsoft.
        let env = |k: &str| if cfg!(debug_assertions) { std::env::var(k).ok() } else { None };
        Self {
            login: env("DE_LOGIN_BASE").unwrap_or_else(|| "https://login.microsoftonline.com".into()),
            graph: env("DE_GRAPH_BASE").unwrap_or_else(|| "https://graph.microsoft.com/beta".into()),
        }
    }
}

fn is_local(url: &str) -> bool {
    url.starts_with("http://127.0.0.1") || url.starts_with("http://localhost")
}

/// Windows system proxy ("Internet Options"), used when no proxy environment variable is set.
fn system_proxy() -> Option<ureq::Proxy> {
    if let Some(p) = ureq::Proxy::try_from_env() {
        return Some(p);
    }
    if !cfg!(windows) {
        return None;
    }
    let out = std::process::Command::new("reg")
        .args(["query", r"HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings"])
        .output()
        .ok()?;
    let text = String::from_utf8_lossy(&out.stdout);
    let value = |name: &str| {
        text.lines().find_map(|l| {
            let mut parts = l.split_whitespace();
            (parts.next()? == name).then(|| parts.skip(1).collect::<Vec<_>>().join(" "))
        })
    };
    if value("ProxyEnable")?.trim() != "0x1" {
        return None;
    }
    let server = value("ProxyServer")?;
    // "host:port" or "http=host:port;https=host:port"
    let chosen = server
        .split(';')
        .find_map(|s| s.strip_prefix("https="))
        .or_else(|| server.split(';').find_map(|s| s.strip_prefix("http=")))
        .unwrap_or(&server)
        .trim()
        .to_string();
    ureq::Proxy::new(&format!("http://{chosen}")).ok()
}

pub fn agent_for(url: &str) -> ureq::Agent {
    let tls = ureq::tls::TlsConfig::builder().root_certs(ureq::tls::RootCerts::PlatformVerifier).build();
    let proxy = if is_local(url) { None } else { system_proxy() };
    ureq::Agent::config_builder()
        .http_status_as_error(false)
        .timeout_global(Some(Duration::from_secs(600)))
        .tls_config(tls)
        .proxy(proxy)
        .user_agent(format!("DeploymentEditor/{}", env!("CARGO_PKG_VERSION")))
        .build()
        .into()
}

fn read_json(resp: &mut ureq::http::Response<ureq::Body>) -> Value {
    let text = resp.body_mut().with_config().limit(50 * 1024 * 1024).read_to_string().unwrap_or_default();
    serde_json::from_str(&text).unwrap_or_else(|_| json!({ "raw": text }))
}

fn aad_error(v: &Value) -> String {
    let desc = v["error_description"].as_str().unwrap_or("");
    let first = desc.lines().next().unwrap_or("").trim();
    match v["error"].as_str() {
        Some(code) if !first.is_empty() => format!("{code}: {first}"),
        Some(code) => code.to_string(),
        None => v["raw"].as_str().unwrap_or("unknown error").chars().take(400).collect(),
    }
}

// ---------------------------------------------------------------------------- authentication

pub const REQUIRED_PERMISSION: &str = "DeviceManagementApps.ReadWrite.All";

/// Payload of a JWT access token (not verified — only used for error messages).
pub fn token_claims(access_token: &str) -> Option<Value> {
    use base64::Engine;
    let payload = access_token.split('.').nth(1)?;
    let bytes = base64::engine::general_purpose::URL_SAFE_NO_PAD.decode(payload.trim_end_matches('=')).ok()?;
    serde_json::from_slice(&bytes).ok()
}

/// Explains a 401 / 403 from Intune using the permissions inside the token.
pub fn permission_hint(auth: &Auth, claims: Option<&Value>, status: u16) -> String {
    let roles: Vec<String> = claims
        .and_then(|c| c["roles"].as_array())
        .map(|a| a.iter().filter_map(|r| r.as_str().map(str::to_string)).collect())
        .unwrap_or_default();
    let scopes: Vec<String> = claims
        .and_then(|c| c["scp"].as_str())
        .map(|s| s.split_whitespace().map(str::to_string).collect())
        .unwrap_or_default();
    let has = |list: &[String]| list.iter().any(|r| r.eq_ignore_ascii_case(REQUIRED_PERMISSION) || r.eq_ignore_ascii_case("DeviceManagementApps.ReadWrite"));
    let tenant = claims.and_then(|c| c["tid"].as_str()).unwrap_or("?");
    let list = |l: &[String]| if l.is_empty() { "none".to_string() } else { l.join(", ") };
    match auth {
        Auth::ClientSecret { .. } if !has(&roles) => {
            let delegated = if has(&scopes) {
                "\nThe permission was added as *Delegated* — the app registration needs it as *Application* permission."
            } else {
                ""
            };
            format!(
                "\n\nThe sign-in worked, but the token has no Intune permission.{delegated}\n\
                 Fix: Entra ID → App registrations → your app → API permissions → Add a permission → Microsoft Graph → \
                 Application permissions → {REQUIRED_PERMISSION}, then click \"Grant admin consent\".\n\
                 Application permissions in the token: {}",
                list(&roles)
            )
        }
        Auth::DeviceCode { .. } if !has(&scopes) => format!(
            "\n\nThe sign-in worked, but the token has no Intune permission ({REQUIRED_PERMISSION}). \
             An administrator has to grant consent for the client app.\nScopes in the token: {}",
            list(&scopes)
        ),
        Auth::ClientSecret { .. } => format!(
            "\n\nThe token contains {REQUIRED_PERMISSION}, but Intune still refused (HTTP {status}). \
             Check that tenant {tenant} has an active Intune license and that this is the tenant you manage. \
             Right after granting consent, it can take a few minutes until Intune accepts the permission."
        ),
        Auth::DeviceCode { .. } => format!(
            "\n\nThe token contains {REQUIRED_PERMISSION}, but Intune still refused (HTTP {status}). \
             Your account needs an Intune role that may manage apps (e.g. Intune Administrator or Application Manager), \
             and tenant {tenant} needs an active Intune license."
        ),
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Auth {
    ClientSecret { tenant: String, client_id: String, secret: String },
    DeviceCode { tenant: String, client_id: String },
}

#[derive(Clone, Debug)]
pub struct Token {
    pub access_token: String,
    pub refresh_token: Option<String>,
    pub expires: Instant,
}

impl Token {
    pub fn valid_for(&self, d: Duration) -> bool {
        self.expires > Instant::now() + d
    }
}

/// Intune wraps its own JSON into "Message" — keep only the readable part.
fn short_intune_message(v: &Value) -> String {
    let inner = v["Message"].as_str().and_then(|m| serde_json::from_str::<Value>(m).ok());
    let text = inner
        .as_ref()
        .and_then(|i| i["Message"].as_str())
        .map(str::to_string)
        .or_else(|| v["Message"].as_str().map(str::to_string));
    match (v["ErrorCode"].as_str(), text) {
        (Some(code), Some(t)) => {
            let t = t.split(" - Url:").next().unwrap_or(&t).trim().to_string();
            format!("{code}: {t}")
        }
        (_, Some(t)) => t,
        _ => v.to_string(),
    }
}

fn parse_token(v: &Value) -> Result<Token> {
    let access = v["access_token"].as_str().ok_or_else(|| anyhow!("Sign-in failed: {}", aad_error(v)))?;
    let secs = v["expires_in"].as_u64().or_else(|| v["expires_in"].as_str().and_then(|s| s.parse().ok())).unwrap_or(3600);
    Ok(Token {
        access_token: access.to_string(),
        refresh_token: v["refresh_token"].as_str().map(str::to_string),
        expires: Instant::now() + Duration::from_secs(secs.saturating_sub(60)),
    })
}

fn token_url(ep: &Endpoints, tenant: &str) -> String {
    format!("{}/{}/oauth2/v2.0/token", ep.login, tenant.trim())
}

pub fn token_client_secret(ep: &Endpoints, tenant: &str, client_id: &str, secret: &str) -> Result<Token> {
    if tenant.trim().is_empty() || client_id.trim().is_empty() || secret.is_empty() {
        bail!("Tenant ID, client ID and client secret are required.");
    }
    let url = token_url(ep, tenant);
    let mut resp = agent_for(&url)
        .post(&url)
        .send_form([
            ("grant_type", "client_credentials"),
            ("client_id", client_id.trim()),
            ("client_secret", secret),
            ("scope", SCOPE_APP),
        ])
        .map_err(|e| anyhow!("Can't reach Microsoft Entra ID: {e}"))?;
    parse_token(&read_json(&mut resp))
}

pub fn token_refresh(ep: &Endpoints, tenant: &str, client_id: &str, refresh_token: &str) -> Result<Token> {
    let url = token_url(ep, tenant);
    let mut resp = agent_for(&url)
        .post(&url)
        .send_form([
            ("grant_type", "refresh_token"),
            ("client_id", client_id.trim()),
            ("refresh_token", refresh_token),
            ("scope", SCOPE_DELEGATED),
        ])
        .map_err(|e| anyhow!("Can't reach Microsoft Entra ID: {e}"))?;
    parse_token(&read_json(&mut resp))
}

#[derive(Clone, Debug)]
pub struct DeviceCodePrompt {
    pub device_code: String,
    pub user_code: String,
    pub verification_uri: String,
    pub interval: u64,
    pub expires: Instant,
}

pub fn device_code_start(ep: &Endpoints, tenant: &str, client_id: &str) -> Result<DeviceCodePrompt> {
    let tenant = if tenant.trim().is_empty() { "organizations" } else { tenant.trim() };
    let url = format!("{}/{}/oauth2/v2.0/devicecode", ep.login, tenant);
    let mut resp = agent_for(&url)
        .post(&url)
        .send_form([("client_id", client_id.trim()), ("scope", SCOPE_DELEGATED)])
        .map_err(|e| anyhow!("Can't reach Microsoft Entra ID: {e}"))?;
    let v = read_json(&mut resp);
    let code = v["device_code"].as_str().ok_or_else(|| anyhow!("Device sign-in failed: {}", aad_error(&v)))?;
    Ok(DeviceCodePrompt {
        device_code: code.to_string(),
        user_code: v["user_code"].as_str().unwrap_or("").to_string(),
        verification_uri: v["verification_uri"].as_str().unwrap_or("https://microsoft.com/devicelogin").to_string(),
        interval: v["interval"].as_u64().unwrap_or(5).max(1),
        expires: Instant::now() + Duration::from_secs(v["expires_in"].as_u64().unwrap_or(900)),
    })
}

/// Wait until the user finished the device sign-in. `cancelled()` is checked between polls.
pub fn device_code_wait(ep: &Endpoints, tenant: &str, client_id: &str, prompt: &DeviceCodePrompt, cancelled: &dyn Fn() -> bool) -> Result<Token> {
    let tenant = if tenant.trim().is_empty() { "organizations" } else { tenant.trim() };
    let url = token_url(ep, tenant);
    let agent = agent_for(&url);
    let mut interval = prompt.interval;
    loop {
        for _ in 0..interval * 10 {
            if cancelled() {
                bail!("Sign-in cancelled.");
            }
            std::thread::sleep(Duration::from_millis(100));
        }
        if Instant::now() > prompt.expires {
            bail!("The sign-in code expired. Please try again.");
        }
        let mut resp = agent
            .post(&url)
            .send_form([
                ("grant_type", "urn:ietf:params:oauth:grant-type:device_code"),
                ("client_id", client_id.trim()),
                ("device_code", prompt.device_code.as_str()),
            ])
            .map_err(|e| anyhow!("Can't reach Microsoft Entra ID: {e}"))?;
        let v = read_json(&mut resp);
        match v["error"].as_str() {
            None => return parse_token(&v),
            Some("authorization_pending") => {}
            Some("slow_down") => interval += 5,
            Some(_) => bail!("Sign-in failed: {}", aad_error(&v)),
        }
    }
}

/// Client secret storage in the Windows Credential Manager (best effort).
pub fn stored_secret(tenant: &str, client_id: &str) -> Option<String> {
    keyring::Entry::new(KEYRING_SERVICE, &format!("{}/{}", tenant.trim(), client_id.trim())).ok()?.get_password().ok()
}

pub fn store_secret(tenant: &str, client_id: &str, secret: Option<&str>) -> Result<()> {
    let entry = keyring::Entry::new(KEYRING_SERVICE, &format!("{}/{}", tenant.trim(), client_id.trim()))?;
    match secret {
        Some(s) => entry.set_password(s)?,
        None => {
            let _ = entry.delete_credential();
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------- graph client

pub struct Graph {
    ep: Endpoints,
    agent: ureq::Agent,
    auth: Auth,
    token: Token,
}

impl Graph {
    pub fn new(ep: Endpoints, auth: Auth, token: Token) -> Self {
        let agent = agent_for(&ep.graph);
        Self { ep, agent, auth, token }
    }

    fn ensure_token(&mut self) -> Result<()> {
        if self.token.valid_for(Duration::from_secs(300)) {
            return Ok(());
        }
        self.token = match &self.auth {
            Auth::ClientSecret { tenant, client_id, secret } => token_client_secret(&self.ep, tenant, client_id, secret)?,
            Auth::DeviceCode { tenant, client_id } => {
                let rt = self.token.refresh_token.clone().ok_or_else(|| anyhow!("The sign-in expired. Please sign in again."))?;
                let t = if tenant.trim().is_empty() { "organizations" } else { tenant.as_str() };
                token_refresh(&self.ep, t, client_id, &rt)?
            }
        };
        Ok(())
    }

    /// Graph request with retries for throttling / transient errors.
    pub fn call(&mut self, method: &str, path: &str, body: Option<&Value>) -> Result<Value> {
        self.ensure_token()?;
        let url = if path.starts_with("http") { path.to_string() } else { format!("{}/{}", self.ep.graph, path.trim_start_matches('/')) };
        let mut attempt = 0;
        let mut renewed = false;
        loop {
            attempt += 1;
            let auth = format!("Bearer {}", self.token.access_token);
            let result = match method {
                "GET" => self.agent.get(&url).header("Authorization", &auth).call(),
                "DELETE" => self.agent.delete(&url).header("Authorization", &auth).call(),
                "POST" => self.agent.post(&url).header("Authorization", &auth).send_json(body.unwrap_or(&json!({}))),
                "PATCH" => self.agent.patch(&url).header("Authorization", &auth).send_json(body.unwrap_or(&json!({}))),
                _ => bail!("Unsupported method {method}"),
            };
            let mut resp = match result {
                Ok(r) => r,
                Err(e) if attempt < 4 => {
                    std::thread::sleep(Duration::from_secs(2 * attempt));
                    let _ = e;
                    continue;
                }
                Err(e) => bail!("Network error calling Microsoft Graph: {e}"),
            };
            let status = resp.status().as_u16();
            if (status == 429 || status == 503 || status == 502 || status == 504) && attempt < 6 {
                let wait = resp
                    .headers()
                    .get("Retry-After")
                    .and_then(|v| v.to_str().ok())
                    .and_then(|v| v.parse::<u64>().ok())
                    .unwrap_or(5 * attempt)
                    .min(60);
                std::thread::sleep(Duration::from_secs(wait));
                continue;
            }
            let v = read_json(&mut resp);
            if (200..300).contains(&status) {
                return Ok(v);
            }
            // A token from before a permission change doesn't contain the new permission:
            // get a fresh one once and retry.
            if status == 401 && !renewed && !path.starts_with("http") {
                renewed = true;
                self.token.expires = Instant::now();
                if self.ensure_token().is_ok() {
                    continue;
                }
            }
            let msg = v["error"]["message"].as_str().map(str::to_string).unwrap_or_else(|| short_intune_message(&v));
            let hint = match status {
                401 | 403 => permission_hint(&self.auth, token_claims(&self.token.access_token).as_ref(), status),
                _ => String::new(),
            };
            bail!("Microsoft Graph returned {status} for {method} {path}:\n{msg}{hint}");
        }
    }

    fn wait_for_state(&mut self, file_path: &str, stage: &str, timeout: Duration) -> Result<Value> {
        let started = Instant::now();
        let mut polls = 0u32;
        loop {
            let v = self.call("GET", file_path, None)?;
            let state = v["uploadState"].as_str().unwrap_or("").to_ascii_lowercase();
            let stage_l = stage.to_ascii_lowercase();
            if state == format!("{stage_l}success") {
                return Ok(v);
            }
            if state == format!("{stage_l}failed") || state == format!("{stage_l}timedout") {
                bail!("Intune reported '{}' while processing the file.", v["uploadState"].as_str().unwrap_or(""));
            }
            if started.elapsed() > timeout {
                bail!("Timed out waiting for Intune ({stage}). Last state: {}", v["uploadState"]);
            }
            polls += 1;
            std::thread::sleep(Duration::from_secs(match polls {
                0..=5 => 1,
                6..=15 => 3,
                _ => 5,
            }));
        }
    }
}

// ---------------------------------------------------------------------------- app description

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MinOs {
    W10_1607,
    W10_21H2,
    W10_22H2,
    W11_21H2,
    W11_22H2,
    W11_23H2,
    W11_24H2,
}

impl MinOs {
    pub const ALL: [MinOs; 7] =
        [MinOs::W10_1607, MinOs::W10_21H2, MinOs::W10_22H2, MinOs::W11_21H2, MinOs::W11_22H2, MinOs::W11_23H2, MinOs::W11_24H2];
    pub fn label(&self) -> &'static str {
        match self {
            MinOs::W10_1607 => "Windows 10 1607",
            MinOs::W10_21H2 => "Windows 10 21H2",
            MinOs::W10_22H2 => "Windows 10 22H2",
            MinOs::W11_21H2 => "Windows 11 21H2",
            MinOs::W11_22H2 => "Windows 11 22H2",
            MinOs::W11_23H2 => "Windows 11 23H2",
            MinOs::W11_24H2 => "Windows 11 24H2",
        }
    }
    /// Value of `minimumSupportedWindowsRelease` (same table as IntuneWin32App).
    pub fn graph(&self) -> &'static str {
        match self {
            MinOs::W10_1607 => "1607",
            MinOs::W10_21H2 => "Windows10_21H2",
            MinOs::W10_22H2 => "Windows10_22H2",
            MinOs::W11_21H2 => "Windows11_21H2",
            MinOs::W11_22H2 => "Windows11_22H2",
            MinOs::W11_23H2 => "Windows11_23H2",
            MinOs::W11_24H2 => "Windows11_24H2",
        }
    }
    pub fn from_key(s: &str) -> Option<MinOs> {
        Self::ALL.into_iter().find(|m| m.graph() == s)
    }
}

#[derive(Clone, Debug)]
pub struct AppSpec {
    pub display_name: String,
    pub description: String,
    pub publisher: String,
    pub version: String,
    pub developer: String,
    pub notes: String,
    pub install_cmd: String,
    pub uninstall_cmd: String,
    pub run_as_system: bool,
    /// allow | basedOnReturnCode | suppress | force
    pub restart_behavior: String,
    pub max_runtime_minutes: u32,
    pub allow_available_uninstall: bool,
    /// x64 | x86 | arm64 | "x64,x86" | "x64,x86,arm64"
    pub architectures: String,
    pub min_os: MinOs,
    pub detection: Vec<Value>,
    pub icon_png: Option<Vec<u8>>,
}

pub fn app_body(spec: &AppSpec, setup_file: &str, include_content_fields: bool) -> Value {
    use base64::Engine;
    let mut body = json!({
        "@odata.type": "#microsoft.graph.win32LobApp",
        "displayName": spec.display_name.trim(),
        "description": if spec.description.trim().is_empty() { spec.display_name.trim() } else { spec.description.trim() },
        "publisher": if spec.publisher.trim().is_empty() { "Unknown" } else { spec.publisher.trim() },
        "displayVersion": spec.version.trim(),
        "developer": spec.developer.trim(),
        "owner": "",
        "notes": spec.notes.trim(),
        "informationUrl": "",
        "privacyInformationUrl": "",
        "isFeatured": false,
        "installCommandLine": spec.install_cmd.trim(),
        "uninstallCommandLine": spec.uninstall_cmd.trim(),
        "installExperience": {
            "runAsAccount": if spec.run_as_system { "system" } else { "user" },
            "deviceRestartBehavior": spec.restart_behavior,
            "maxRunTimeInMinutes": spec.max_runtime_minutes.clamp(1, 1440),
        },
        "allowAvailableUninstall": spec.allow_available_uninstall,
        "allowedArchitectures": spec.architectures,
        "applicableArchitectures": "none",
        "minimumSupportedWindowsRelease": spec.min_os.graph(),
        "msiInformation": Value::Null,
        "runAs32bit": false,
        "returnCodes": [
            { "returnCode": 0, "type": "success" },
            { "returnCode": 1707, "type": "success" },
            { "returnCode": 3010, "type": "softReboot" },
            { "returnCode": 1641, "type": "hardReboot" },
            { "returnCode": 1618, "type": "retry" },
        ],
        "rules": spec.detection,
    });
    if include_content_fields {
        body["fileName"] = json!(crate::intunewin::ENCRYPTED_FILE_NAME);
        body["setupFilePath"] = json!(setup_file);
    }
    if let Some(png) = &spec.icon_png {
        body["largeIcon"] = json!({
            "@odata.type": "#microsoft.graph.mimeContent",
            "type": "image/png",
            "value": base64::engine::general_purpose::STANDARD.encode(png),
        });
    }
    body
}

#[derive(Clone, Debug, PartialEq)]
pub struct ExistingApp {
    pub id: String,
    pub display_name: String,
    pub version: String,
    pub publisher: String,
}

/// Win32 apps whose display name contains `name` (case-insensitive; filtered locally).
pub fn find_apps(graph: &mut Graph, name: &str) -> Result<Vec<ExistingApp>> {
    let filter = "isof('microsoft.graph.win32LobApp')";
    let mut url = format!("deviceAppManagement/mobileApps?$filter={}&$select=id,displayName,displayVersion,publisher&$top=999", urlencode(filter));
    let needle = name.trim().to_lowercase();
    let mut out = Vec::new();
    for _ in 0..20 {
        let v = graph.call("GET", &url, None)?;
        for a in v["value"].as_array().cloned().unwrap_or_default() {
            let dn = a["displayName"].as_str().unwrap_or("").to_string();
            if needle.is_empty() || dn.to_lowercase().contains(&needle) {
                out.push(ExistingApp {
                    id: a["id"].as_str().unwrap_or("").to_string(),
                    display_name: dn,
                    version: a["displayVersion"].as_str().unwrap_or("").to_string(),
                    publisher: a["publisher"].as_str().unwrap_or("").to_string(),
                });
            }
        }
        match v["@odata.nextLink"].as_str() {
            Some(next) => url = next.to_string(),
            None => break,
        }
    }
    out.sort_by(|a, b| a.display_name.to_lowercase().cmp(&b.display_name.to_lowercase()));
    Ok(out)
}

pub fn urlencode(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => out.push(b as char),
            _ => out.push_str(&format!("%{b:02X}")),
        }
    }
    out
}

/// Portal link for an app.
pub fn portal_url(app_id: &str) -> String {
    format!("https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/{app_id}")
}

/// Upload the encrypted package; creates a new app or a new content version of `existing`.
/// `progress(fraction, text)`. Returns the app id.
pub fn upload(
    graph: &mut Graph,
    spec: &AppSpec,
    existing: Option<&str>,
    pkg: &EncryptedPackage,
    progress: &mut dyn FnMut(f32, String),
    cancelled: &dyn Fn() -> bool,
) -> Result<String> {
    // 1. app
    let app_id = match existing {
        None => {
            progress(0.02, "Creating the Win32 app in Intune".into());
            let v = graph.call("POST", "deviceAppManagement/mobileApps", Some(&app_body(spec, &pkg.setup_file, true)))?;
            v["id"].as_str().ok_or_else(|| anyhow!("Intune did not return an app id"))?.to_string()
        }
        Some(id) => {
            progress(0.02, "Updating the app properties".into());
            graph.call("PATCH", &format!("deviceAppManagement/mobileApps/{id}"), Some(&app_body(spec, &pkg.setup_file, true)))?;
            id.to_string()
        }
    };
    let base = format!("deviceAppManagement/mobileApps/{app_id}/microsoft.graph.win32LobApp");

    // 2. content version + file
    progress(0.04, "Creating a content version".into());
    let cv = graph.call("POST", &format!("{base}/contentVersions"), Some(&json!({})))?;
    let cv_id = cv["id"].as_str().ok_or_else(|| anyhow!("No content version id returned"))?.to_string();
    let file_body = json!({
        "@odata.type": "#microsoft.graph.mobileAppContentFile",
        "name": crate::intunewin::ENCRYPTED_FILE_NAME,
        "size": pkg.unencrypted_size,
        "sizeEncrypted": pkg.encrypted_size,
        "manifest": Value::Null,
        "isDependency": false,
    });
    let f = graph.call("POST", &format!("{base}/contentVersions/{cv_id}/files"), Some(&file_body))?;
    let file_id = f["id"].as_str().ok_or_else(|| anyhow!("No content file id returned"))?.to_string();
    let file_path = format!("{base}/contentVersions/{cv_id}/files/{file_id}");

    progress(0.06, "Waiting for the Azure Storage upload location".into());
    let state = graph.wait_for_state(&file_path, "azureStorageUriRequest", Duration::from_secs(600))?;
    let mut sas = state["azureStorageUri"].as_str().ok_or_else(|| anyhow!("No azureStorageUri returned"))?.to_string();

    // 3. blocks
    let total = pkg.encrypted_size as usize;
    let chunks = total.div_ceil(CHUNK_SIZE).max(1);
    let mut file = std::fs::File::open(&pkg.path).context("Opening the encrypted package")?;
    let blob = agent_for(&sas);
    let mut sas_age = Instant::now();
    let mut ids = Vec::with_capacity(chunks);
    let mut buf = vec![0u8; CHUNK_SIZE];
    for i in 0..chunks {
        if cancelled() {
            bail!(
                "Upload cancelled. {}",
                if existing.is_none() { "The app was created in Intune without content — upload again to complete it, or delete it." } else { "The app keeps its previous content." }
            );
        }
        if sas_age.elapsed() > Duration::from_secs(7 * 60) {
            progress(0.08 + 0.82 * i as f32 / chunks as f32, "Renewing the upload URL".into());
            graph.call("POST", &format!("{file_path}/renewUpload"), Some(&json!({})))?;
            let st = graph.wait_for_state(&file_path, "azureStorageUriRenewal", Duration::from_secs(120))?;
            sas = st["azureStorageUri"].as_str().unwrap_or(&sas).to_string();
            sas_age = Instant::now();
        }
        let len = std::cmp::min(CHUNK_SIZE, total - i * CHUNK_SIZE);
        use std::io::Read;
        file.read_exact(&mut buf[..len])?;
        use base64::Engine;
        let id = base64::engine::general_purpose::STANDARD.encode(format!("{i:04}"));
        let url = format!("{sas}&comp=block&blockid={}", urlencode(&id));
        let mut ok = false;
        let mut last_err = String::new();
        for attempt in 0..6u64 {
            match blob.put(&url).header("x-ms-blob-type", "BlockBlob").send(&buf[..len]) {
                Ok(r) if r.status().is_success() => {
                    ok = true;
                    break;
                }
                Ok(mut r) => last_err = format!("HTTP {}: {}", r.status(), r.body_mut().read_to_string().unwrap_or_default()),
                Err(e) => last_err = e.to_string(),
            }
            std::thread::sleep(Duration::from_secs(3 + attempt * 4));
        }
        if !ok {
            bail!("Uploading part {} of {} failed: {last_err}", i + 1, chunks);
        }
        ids.push(id);
        let done = ((i + 1) * CHUNK_SIZE).min(total);
        progress(
            0.08 + 0.82 * (i + 1) as f32 / chunks as f32,
            format!("Uploading {:.1} of {:.1} MB", done as f64 / 1_048_576.0, total as f64 / 1_048_576.0),
        );
    }
    let mut xml = String::from("<?xml version=\"1.0\" encoding=\"utf-8\"?><BlockList>");
    for id in &ids {
        xml.push_str(&format!("<Latest>{id}</Latest>"));
    }
    xml.push_str("</BlockList>");
    let r = blob
        .put(&format!("{sas}&comp=blocklist"))
        .header("Content-Type", "text/plain; charset=UTF-8")
        .send(xml.as_bytes())
        .map_err(|e| anyhow!("Finalizing the upload failed: {e}"))?;
    if !r.status().is_success() {
        bail!("Finalizing the upload failed (HTTP {}).", r.status());
    }

    // 4. commit
    progress(0.92, "Committing the file (Intune decrypts and verifies it)".into());
    graph.call("POST", &format!("{file_path}/commit"), Some(&json!({ "fileEncryptionInfo": pkg.info.to_graph() })))?;
    graph.wait_for_state(&file_path, "commitFile", Duration::from_secs(1800))?;

    progress(0.98, "Publishing the new content version".into());
    graph.call(
        "PATCH",
        &format!("deviceAppManagement/mobileApps/{app_id}"),
        Some(&json!({ "@odata.type": "#microsoft.graph.win32LobApp", "committedContentVersion": cv_id })),
    )?;
    progress(1.0, "Upload complete".into());
    Ok(app_id)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fake_jwt(claims: serde_json::Value) -> String {
        use base64::Engine;
        let e = base64::engine::general_purpose::URL_SAFE_NO_PAD;
        format!("{}.{}.sig", e.encode(br#"{"alg":"none"}"#), e.encode(claims.to_string()))
    }

    #[test]
    fn permission_hints() {
        let app = Auth::ClientSecret { tenant: "t".into(), client_id: "c".into(), secret: "s".into() };
        let user = Auth::DeviceCode { tenant: "t".into(), client_id: "c".into() };

        // App registration without any application permission
        let c = token_claims(&fake_jwt(json!({"tid": "abc", "roles": []}))).unwrap();
        let h = permission_hint(&app, Some(&c), 401);
        assert!(h.contains("token has no Intune permission") && h.contains("Grant admin consent") && h.contains("none"), "{h}");
        assert!(!h.contains("Delegated"));

        // Added as delegated instead of application
        let c = token_claims(&fake_jwt(json!({"tid": "abc", "scp": "DeviceManagementApps.ReadWrite.All"}))).unwrap();
        assert!(permission_hint(&app, Some(&c), 401).contains("*Delegated*"));

        // Permission present: license / tenant hint
        let c = token_claims(&fake_jwt(json!({"tid": "abc", "roles": ["DeviceManagementApps.ReadWrite.All"]}))).unwrap();
        let h = permission_hint(&app, Some(&c), 401);
        assert!(h.contains("Intune license") && h.contains("abc"), "{h}");

        // Delegated sign-in
        let c = token_claims(&fake_jwt(json!({"tid": "abc", "scp": "openid profile"}))).unwrap();
        assert!(permission_hint(&user, Some(&c), 403).contains("openid, profile"));
        let c = token_claims(&fake_jwt(json!({"tid": "abc", "scp": "DeviceManagementApps.ReadWrite.All openid"}))).unwrap();
        assert!(permission_hint(&user, Some(&c), 403).contains("Intune role"));

        // Not a JWT
        assert!(token_claims("opaque-token").is_none());
        assert!(permission_hint(&app, None, 401).contains("no Intune permission"));
    }

    #[test]
    fn intune_error_message_is_shortened() {
        let v: serde_json::Value = serde_json::from_str(r#"{"ErrorCode":"Forbidden","Message":"{\r\n  \"_version\": 3,\r\n  \"Message\": \"An error has occurred - Operation ID (for customer support): 00000000-0000-0000-0000-000000000000 - Activity ID: 96cb80e2 - Url: https://proxy.msub09.manage.microsoft.com/x\",\r\n  \"RetryAfter\": null\r\n}","Target":null}"#).unwrap();
        let m = short_intune_message(&v);
        assert_eq!(m, "Forbidden: An error has occurred - Operation ID (for customer support): 00000000-0000-0000-0000-000000000000 - Activity ID: 96cb80e2");
    }

    #[test]
    fn body_has_expected_fields() {
        let spec = AppSpec {
            display_name: "Igor Pavlov 7-Zip 25.01 x64".into(),
            description: String::new(),
            publisher: "Igor Pavlov".into(),
            version: "25.01".into(),
            developer: String::new(),
            notes: String::new(),
            install_cmd: "Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent".into(),
            uninstall_cmd: "Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent".into(),
            run_as_system: true,
            restart_behavior: "suppress".into(),
            max_runtime_minutes: 60,
            allow_available_uninstall: true,
            architectures: "x64".into(),
            min_os: MinOs::W11_21H2,
            detection: vec![json!({"ruleType": "detection"})],
            icon_png: Some(vec![1, 2, 3]),
        };
        let b = app_body(&spec, "Invoke-AppDeployToolkit.exe", true);
        assert_eq!(b["@odata.type"], "#microsoft.graph.win32LobApp");
        assert_eq!(b["description"], "Igor Pavlov 7-Zip 25.01 x64", "empty description falls back to the name");
        assert_eq!(b["minimumSupportedWindowsRelease"], "Windows11_21H2");
        assert_eq!(b["allowedArchitectures"], "x64");
        assert_eq!(b["applicableArchitectures"], "none");
        assert_eq!(b["installExperience"]["deviceRestartBehavior"], "suppress");
        assert_eq!(b["setupFilePath"], "Invoke-AppDeployToolkit.exe");
        assert_eq!(b["fileName"], "IntunePackage.intunewin");
        assert_eq!(b["largeIcon"]["value"], "AQID");
        assert_eq!(b["returnCodes"].as_array().unwrap().len(), 5);
        assert_eq!(urlencode("MDAwMA=="), "MDAwMA%3D%3D");
        assert_eq!(MinOs::from_key("Windows10_22H2"), Some(MinOs::W10_22H2));
    }

    /// End-to-end against `tests/mock_graph.py` (started by the test script).
    #[test]
    #[ignore]
    fn upload_against_mock_graph() {
        let root = std::env::temp_dir().join("de-mock-upload");
        let _ = std::fs::remove_dir_all(&root);
        let src = root.join("src");
        std::fs::create_dir_all(src.join("Files")).unwrap();
        std::fs::write(src.join("Invoke-AppDeployToolkit.exe"), b"MZ").unwrap();
        // pseudo-random (incompressible) data, so the encrypted file spans 3 blocks
        let mut x: u64 = 0x9E37_79B9_7F4A_7C15;
        let big: Vec<u8> = (0..(13 * 1024 * 1024))
            .map(|_| {
                x ^= x << 13;
                x ^= x >> 7;
                x ^= x << 17;
                x as u8
            })
            .collect();
        std::fs::write(src.join("Files").join("setup.bin"), big).unwrap();
        let pkg = crate::intunewin::build_encrypted(&src, "Invoke-AppDeployToolkit.exe", &root.join("work"), &mut |_, _| {}).unwrap();

        let ep = Endpoints::default();
        assert!(ep.graph.starts_with("http://127.0.0.1"), "set DE_GRAPH_BASE / DE_LOGIN_BASE to the mock server");
        let token = token_client_secret(&ep, "tenant-1", "client-1", "secret-1").unwrap();
        let auth = Auth::ClientSecret { tenant: "tenant-1".into(), client_id: "client-1".into(), secret: "secret-1".into() };
        let mut graph = Graph::new(ep.clone(), auth, token);
        assert!(token_client_secret(&ep, "tenant-1", "client-1", "wrong").unwrap_err().to_string().contains("invalid_client"));

        // App registration without Intune permission: readable hint, one retry with a fresh token
        let t = token_client_secret(&ep, "tenant-1", "client-noperm", "secret-noperm").unwrap();
        let noperm = Auth::ClientSecret { tenant: "tenant-1".into(), client_id: "client-noperm".into(), secret: "secret-noperm".into() };
        let mut g = Graph::new(ep.clone(), noperm, t);
        let err = g.call("POST", "deviceAppManagement/mobileApps", Some(&json!({}))).unwrap_err().to_string();
        eprintln!("{err}");
        assert!(err.contains("returned 401"), "{err}");
        assert!(err.contains("Forbidden: An error has occurred"), "{err}");
        assert!(!err.contains("_version"), "raw Intune JSON should be shortened: {err}");
        assert!(err.contains("token has no Intune permission") && err.contains("Grant admin consent"), "{err}");

        let spec = AppSpec {
            display_name: "Mock App 1.0".into(),
            description: "desc".into(),
            publisher: "Vendor".into(),
            version: "1.0".into(),
            developer: String::new(),
            notes: String::new(),
            install_cmd: "install".into(),
            uninstall_cmd: "uninstall".into(),
            run_as_system: true,
            restart_behavior: "suppress".into(),
            max_runtime_minutes: 60,
            allow_available_uninstall: true,
            architectures: "x64".into(),
            min_os: MinOs::W11_21H2,
            detection: vec![crate::detection::DetectionRule::Msi {
                product_code: "{23170F69-40C1-2702-2501-000001000000}".into(),
                version_operator: crate::detection::Operator::NotConfigured,
                product_version: String::new(),
            }
            .to_graph()],
            icon_png: None,
        };
        let mut steps = Vec::new();
        let id = upload(&mut graph, &spec, None, &pkg, &mut |f, t| steps.push((f, t)), &|| false).unwrap();
        assert_eq!(id, "app-1");
        assert!(steps.iter().any(|(_, t)| t.starts_with("Uploading")));
        let found = find_apps(&mut graph, "mock").unwrap();
        assert_eq!(found.len(), 1);
        // second upload = new content version of the same app
        let id2 = upload(&mut graph, &spec, Some(&id), &pkg, &mut |_, _| {}, &|| false).unwrap();
        assert_eq!(id2, id);
        std::fs::write(root.join("encrypted-path.txt"), pkg.path.to_string_lossy().as_bytes()).unwrap();
    }
}

#[cfg(test)]
mod live_tests {
    /// Talks to the real Microsoft login endpoint with a non-existing tenant.
    #[test]
    #[ignore]
    fn real_login_reports_readable_error() {
        let ep = super::Endpoints::default();
        let err = super::token_client_secret(&ep, "00000000-0000-0000-0000-000000000001", "11111111-1111-1111-1111-111111111111", "x")
            .unwrap_err()
            .to_string();
        eprintln!("{err}");
        assert!(err.contains("AADSTS"), "{err}");
    }
}
