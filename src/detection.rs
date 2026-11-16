//! Intune detection rules for a project: model, persistence, suggestions and
//! conversion to Microsoft Graph `win32LobAppRule` JSON.
//!
//! The rule is stored as JSON in the project's `Settings` table
//! (`Intune_DetectionRule`), so older editor versions simply ignore it.

use serde::{Deserialize, Serialize};
use serde_json::{json, Value};

use crate::project::{keys, Action, DeploymentType, Project};

pub const SETTING_KEY: &str = "Intune_DetectionRule";
/// MSI ProductCode remembered from an MSI / WinGet import.
pub const PRODUCT_CODE_KEY: &str = "App_ProductCode";

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub enum Operator {
    #[default]
    NotConfigured,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

impl Operator {
    pub const ALL: [Operator; 6] = [
        Operator::Equal,
        Operator::NotEqual,
        Operator::GreaterThan,
        Operator::GreaterThanOrEqual,
        Operator::LessThan,
        Operator::LessThanOrEqual,
    ];
    pub fn graph(&self) -> &'static str {
        match self {
            Operator::NotConfigured => "notConfigured",
            Operator::Equal => "equal",
            Operator::NotEqual => "notEqual",
            Operator::GreaterThan => "greaterThan",
            Operator::GreaterThanOrEqual => "greaterThanOrEqual",
            Operator::LessThan => "lessThan",
            Operator::LessThanOrEqual => "lessThanOrEqual",
        }
    }
    pub fn symbol(&self) -> &'static str {
        match self {
            Operator::NotConfigured => "any",
            Operator::Equal => "=",
            Operator::NotEqual => "≠",
            Operator::GreaterThan => ">",
            Operator::GreaterThanOrEqual => "≥",
            Operator::LessThan => "<",
            Operator::LessThanOrEqual => "≤",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub enum FileCheck {
    #[default]
    Exists,
    DoesNotExist,
    Version,
    SizeInMB,
}

impl FileCheck {
    pub const ALL: [FileCheck; 4] = [FileCheck::Exists, FileCheck::DoesNotExist, FileCheck::Version, FileCheck::SizeInMB];
    pub fn graph(&self) -> &'static str {
        match self {
            FileCheck::Exists => "exists",
            FileCheck::DoesNotExist => "doesNotExist",
            FileCheck::Version => "version",
            FileCheck::SizeInMB => "sizeInMB",
        }
    }
    pub fn label(&self) -> &'static str {
        match self {
            FileCheck::Exists => "File or folder exists",
            FileCheck::DoesNotExist => "File or folder does not exist",
            FileCheck::Version => "File version",
            FileCheck::SizeInMB => "Size in MB",
        }
    }
    pub fn needs_value(&self) -> bool {
        matches!(self, FileCheck::Version | FileCheck::SizeInMB)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub enum RegistryCheck {
    #[default]
    Exists,
    DoesNotExist,
    String,
    Integer,
    Version,
}

impl RegistryCheck {
    pub const ALL: [RegistryCheck; 5] =
        [RegistryCheck::Exists, RegistryCheck::DoesNotExist, RegistryCheck::String, RegistryCheck::Integer, RegistryCheck::Version];
    pub fn graph(&self) -> &'static str {
        match self {
            RegistryCheck::Exists => "exists",
            RegistryCheck::DoesNotExist => "doesNotExist",
            RegistryCheck::String => "string",
            RegistryCheck::Integer => "integer",
            RegistryCheck::Version => "version",
        }
    }
    pub fn label(&self) -> &'static str {
        match self {
            RegistryCheck::Exists => "Key / value exists",
            RegistryCheck::DoesNotExist => "Key / value does not exist",
            RegistryCheck::String => "String comparison",
            RegistryCheck::Integer => "Integer comparison",
            RegistryCheck::Version => "Version comparison",
        }
    }
    pub fn needs_value(&self) -> bool {
        matches!(self, RegistryCheck::String | RegistryCheck::Integer | RegistryCheck::Version)
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum DetectionRule {
    #[serde(rename_all = "camelCase")]
    Msi { product_code: String, version_operator: Operator, product_version: String },
    #[serde(rename_all = "camelCase")]
    File { path: String, file_or_folder: String, check: FileCheck, operator: Operator, value: String, check_32bit: bool },
    #[serde(rename_all = "camelCase")]
    Registry { key_path: String, value_name: String, check: RegistryCheck, operator: Operator, value: String, check_32bit: bool },
    #[serde(rename_all = "camelCase")]
    Script { script: String, run_as_32bit: bool, enforce_signature_check: bool },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RuleKind {
    Msi,
    File,
    Registry,
    Script,
}

impl RuleKind {
    pub const ALL: [RuleKind; 4] = [RuleKind::Msi, RuleKind::File, RuleKind::Registry, RuleKind::Script];
    pub fn label(&self) -> &'static str {
        match self {
            RuleKind::Msi => "MSI",
            RuleKind::File => "File",
            RuleKind::Registry => "Registry",
            RuleKind::Script => "Script",
        }
    }
    pub fn empty(&self) -> DetectionRule {
        match self {
            RuleKind::Msi => DetectionRule::Msi {
                product_code: String::new(),
                version_operator: Operator::NotConfigured,
                product_version: String::new(),
            },
            RuleKind::File => DetectionRule::File {
                path: "%ProgramFiles%\\".into(),
                file_or_folder: String::new(),
                check: FileCheck::Exists,
                operator: Operator::NotConfigured,
                value: String::new(),
                check_32bit: false,
            },
            RuleKind::Registry => DetectionRule::Registry {
                key_path: "HKEY_LOCAL_MACHINE\\SOFTWARE\\".into(),
                value_name: String::new(),
                check: RegistryCheck::Exists,
                operator: Operator::NotConfigured,
                value: String::new(),
                check_32bit: false,
            },
            RuleKind::Script => DetectionRule::Script { script: String::new(), run_as_32bit: false, enforce_signature_check: false },
        }
    }
}

pub fn is_guid(s: &str) -> bool {
    let t = s.trim().trim_start_matches('{').trim_end_matches('}');
    let parts: Vec<&str> = t.split('-').collect();
    parts.len() == 5
        && [8, 4, 4, 4, 12].iter().zip(&parts).all(|(n, p)| p.len() == *n && p.chars().all(|c| c.is_ascii_hexdigit()))
}

/// `{GUID}` in upper case, as Intune shows it.
pub fn normalize_guid(s: &str) -> String {
    let t = s.trim().trim_start_matches('{').trim_end_matches('}').to_ascii_uppercase();
    format!("{{{t}}}")
}

fn looks_like_version(s: &str) -> bool {
    let s = s.trim();
    !s.is_empty() && s.split('.').all(|p| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit())) && s.split('.').count() <= 4
}

impl DetectionRule {
    pub fn kind(&self) -> RuleKind {
        match self {
            DetectionRule::Msi { .. } => RuleKind::Msi,
            DetectionRule::File { .. } => RuleKind::File,
            DetectionRule::Registry { .. } => RuleKind::Registry,
            DetectionRule::Script { .. } => RuleKind::Script,
        }
    }

    pub fn load(project: &Project) -> Option<DetectionRule> {
        let raw = project.setting(SETTING_KEY);
        if raw.trim().is_empty() {
            return None;
        }
        serde_json::from_str(raw).ok()
    }

    pub fn save(project: &mut Project, rule: Option<&DetectionRule>) -> anyhow::Result<()> {
        let raw = match rule {
            Some(r) => serde_json::to_string(r)?,
            None => String::new(),
        };
        project.set_setting(SETTING_KEY, &raw)
    }

    /// Problems that would make Intune reject the rule or never detect the app.
    pub fn problems(&self) -> Vec<String> {
        let mut out = Vec::new();
        match self {
            DetectionRule::Msi { product_code, version_operator, product_version } => {
                if !is_guid(product_code) {
                    out.push("The MSI product code must be a GUID like {23170F69-40C1-2702-2501-000001000000}.".into());
                }
                if *version_operator != Operator::NotConfigured && product_version.trim().is_empty() {
                    out.push("Enter the MSI product version to compare with.".into());
                }
            }
            DetectionRule::File { path, file_or_folder, check, operator, value, .. } => {
                if path.trim().is_empty() {
                    out.push("Enter the folder path of the file.".into());
                }
                if file_or_folder.trim().is_empty() {
                    out.push("Enter the file or folder name.".into());
                }
                if path.contains('$') {
                    out.push("Use environment variables like %ProgramFiles% — PowerShell variables are not resolved by Intune.".into());
                }
                if check.needs_value() {
                    if *operator == Operator::NotConfigured {
                        out.push("Choose an operator for the comparison.".into());
                    }
                    if value.trim().is_empty() {
                        out.push("Enter the value to compare with.".into());
                    } else if *check == FileCheck::Version && !looks_like_version(value) {
                        out.push("The version must look like 1.2.3.4.".into());
                    }
                }
            }
            DetectionRule::Registry { key_path, check, operator, value, .. } => {
                let k = key_path.trim().to_ascii_uppercase();
                if !(k.starts_with("HKEY_LOCAL_MACHINE\\") || k.starts_with("HKEY_CURRENT_USER\\") || k.starts_with("HKEY_CLASSES_ROOT\\")) {
                    out.push("The key path must start with HKEY_LOCAL_MACHINE\\ (or HKEY_CURRENT_USER\\).".into());
                }
                if check.needs_value() {
                    if *operator == Operator::NotConfigured {
                        out.push("Choose an operator for the comparison.".into());
                    }
                    if value.trim().is_empty() {
                        out.push("Enter the value to compare with.".into());
                    }
                }
            }
            DetectionRule::Script { script, .. } => {
                if script.trim().is_empty() {
                    out.push("The detection script is empty.".into());
                }
            }
        }
        out
    }

    /// One line for the UI / PDF.
    pub fn summary(&self) -> String {
        match self {
            DetectionRule::Msi { product_code, version_operator, product_version } => {
                if *version_operator == Operator::NotConfigured {
                    format!("MSI product code {product_code} is installed")
                } else {
                    format!("MSI product code {product_code}, version {} {product_version}", version_operator.symbol())
                }
            }
            DetectionRule::File { path, file_or_folder, check, operator, value, check_32bit } => {
                let full = format!("{}\\{}", path.trim_end_matches('\\'), file_or_folder);
                let what = match check {
                    FileCheck::Exists => format!("{full} exists"),
                    FileCheck::DoesNotExist => format!("{full} does not exist"),
                    FileCheck::Version => format!("{full} version {} {value}", operator.symbol()),
                    FileCheck::SizeInMB => format!("{full} size {} {value} MB", operator.symbol()),
                };
                if *check_32bit { format!("{what} (32-bit view)") } else { what }
            }
            DetectionRule::Registry { key_path, value_name, check, operator, value, check_32bit } => {
                let target = if value_name.is_empty() { key_path.clone() } else { format!("{key_path} → {value_name}") };
                let what = match check {
                    RegistryCheck::Exists => format!("{target} exists"),
                    RegistryCheck::DoesNotExist => format!("{target} does not exist"),
                    _ => format!("{target} {} {value}", operator.symbol()),
                };
                if *check_32bit { format!("{what} (32-bit view)") } else { what }
            }
            DetectionRule::Script { script, .. } => {
                let lines = script.lines().filter(|l| !l.trim().is_empty()).count();
                format!("PowerShell detection script ({lines} lines)")
            }
        }
    }

    /// `win32LobAppRule` JSON (Graph beta `rules` collection, ruleType = detection).
    pub fn to_graph(&self) -> Value {
        match self {
            DetectionRule::Msi { product_code, version_operator, product_version } => json!({
                "@odata.type": "#microsoft.graph.win32LobAppProductCodeRule",
                "ruleType": "detection",
                "productCode": normalize_guid(product_code),
                "productVersionOperator": version_operator.graph(),
                "productVersion": if *version_operator == Operator::NotConfigured { Value::Null } else { Value::String(product_version.trim().to_string()) },
            }),
            DetectionRule::File { path, file_or_folder, check, operator, value, check_32bit } => json!({
                "@odata.type": "#microsoft.graph.win32LobAppFileSystemRule",
                "ruleType": "detection",
                "path": path.trim().trim_end_matches('\\'),
                "fileOrFolderName": file_or_folder.trim(),
                "check32BitOn64System": check_32bit,
                "operationType": check.graph(),
                "operator": if check.needs_value() { operator.graph() } else { "notConfigured" },
                "comparisonValue": if check.needs_value() { Value::String(value.trim().to_string()) } else { Value::Null },
            }),
            DetectionRule::Registry { key_path, value_name, check, operator, value, check_32bit } => json!({
                "@odata.type": "#microsoft.graph.win32LobAppRegistryRule",
                "ruleType": "detection",
                "check32BitOn64System": check_32bit,
                "keyPath": key_path.trim().trim_end_matches('\\'),
                "valueName": if value_name.trim().is_empty() { Value::Null } else { Value::String(value_name.trim().to_string()) },
                "operationType": check.graph(),
                "operator": if check.needs_value() { operator.graph() } else { "notConfigured" },
                "comparisonValue": if check.needs_value() { Value::String(value.trim().to_string()) } else { Value::Null },
            }),
            DetectionRule::Script { script, run_as_32bit, enforce_signature_check } => {
                use base64::Engine;
                // Intune expects the script as base64 of the UTF-8 text.
                let content = base64::engine::general_purpose::STANDARD.encode(crate::plugins::to_crlf(script).as_bytes());
                json!({
                    "@odata.type": "#microsoft.graph.win32LobAppPowerShellScriptRule",
                    "ruleType": "detection",
                    "displayName": Value::Null,
                    "enforceSignatureCheck": enforce_signature_check,
                    "runAs32Bit": run_as_32bit,
                    "runAsAccount": Value::Null,
                    "scriptContent": content,
                    "operationType": "notConfigured",
                    "operator": "notConfigured",
                    "comparisonValue": Value::Null,
                })
            }
        }
    }
}

/// Translate PSADT / PowerShell path variables into Intune environment variables.
/// Returns `None` if the path still contains a variable Intune cannot resolve.
pub fn to_intune_path(ps_path: &str) -> Option<String> {
    let mut p = ps_path.trim().trim_matches('"').trim_matches('\'').to_string();
    // longest names first
    let map = [
        ("$envProgramFilesX86", "%ProgramFiles(x86)%"),
        ("${envProgramFilesX86}", "%ProgramFiles(x86)%"),
        ("$envCommonProgramFilesX86", "%CommonProgramFiles(x86)%"),
        ("$envCommonProgramFiles", "%CommonProgramFiles%"),
        ("$envProgramFiles", "%ProgramFiles%"),
        ("${envProgramFiles}", "%ProgramFiles%"),
        ("$env:ProgramFiles(x86)", "%ProgramFiles(x86)%"),
        ("${env:ProgramFiles(x86)}", "%ProgramFiles(x86)%"),
        ("$env:ProgramFiles", "%ProgramFiles%"),
        ("$envProgramData", "%ProgramData%"),
        ("$env:ProgramData", "%ProgramData%"),
        ("$envSystemRoot", "%SystemRoot%"),
        ("$envWinDir", "%windir%"),
        ("$env:SystemRoot", "%SystemRoot%"),
        ("$env:windir", "%windir%"),
        ("$envSystemDrive", "%SystemDrive%"),
        ("$env:SystemDrive", "%SystemDrive%"),
    ];
    for (from, to) in map {
        // case-insensitive replace
        while let Some(pos) = p.to_ascii_lowercase().find(&from.to_ascii_lowercase()) {
            p.replace_range(pos..pos + from.len(), to);
        }
    }
    if p.contains('$') || p.is_empty() {
        None
    } else {
        Some(p)
    }
}

/// Convert `HKLM:\...`, `Registry::HKEY_...` or `HKLM\...` to `HKEY_LOCAL_MACHINE\...`.
pub fn to_intune_registry_path(path: &str) -> Option<String> {
    let p = path.trim().trim_matches('"').trim_matches('\'');
    let p = p.strip_prefix("Registry::").or_else(|| p.strip_prefix("Microsoft.PowerShell.Core\\Registry::")).unwrap_or(p);
    let upper = p.to_ascii_uppercase();
    let (hive, rest) = if let Some(r) = upper.strip_prefix("HKLM:\\").or_else(|| upper.strip_prefix("HKLM\\")) {
        ("HKEY_LOCAL_MACHINE", &p[p.len() - r.len()..])
    } else if let Some(r) = upper.strip_prefix("HKCU:\\").or_else(|| upper.strip_prefix("HKCU\\")) {
        ("HKEY_CURRENT_USER", &p[p.len() - r.len()..])
    } else if let Some(r) = upper.strip_prefix("HKEY_LOCAL_MACHINE\\") {
        ("HKEY_LOCAL_MACHINE", &p[p.len() - r.len()..])
    } else if let Some(r) = upper.strip_prefix("HKEY_CURRENT_USER\\") {
        ("HKEY_CURRENT_USER", &p[p.len() - r.len()..])
    } else {
        return None;
    };
    if rest.contains('$') {
        return None;
    }
    Some(format!("{hive}\\{}", rest.trim_end_matches('\\')))
}

fn ps_single_quote(s: &str) -> String {
    format!("'{}'", s.replace('\'', "''"))
}

/// Generic script: looks for the app in the Uninstall registry keys by display name
/// and requires at least the packaged version.
pub fn uninstall_registry_script(app_name: &str, min_version: &str) -> String {
    format!(
        "# Detection: application registered in Add/Remove Programs with at least the packaged version\r\n\
$name = {}\r\n\
$minVersion = {}\r\n\
$paths = @(\r\n\
    'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',\r\n\
    'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'\r\n\
)\r\n\
$app = Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue |\r\n\
    Where-Object {{ $_.DisplayName -like \"*$name*\" }} |\r\n\
    Select-Object -First 1\r\n\
if ($app) {{\r\n\
    if (-not $minVersion) {{ Write-Output 'Installed'; exit 0 }}\r\n\
    try {{\r\n\
        if ([version]$app.DisplayVersion -ge [version]$minVersion) {{ Write-Output 'Installed'; exit 0 }}\r\n\
    }} catch {{\r\n\
        if ($app.DisplayVersion -eq $minVersion) {{ Write-Output 'Installed'; exit 0 }}\r\n\
    }}\r\n\
}}\r\n\
exit 1\r\n",
        ps_single_quote(app_name),
        ps_single_quote(min_version)
    )
}

#[derive(Clone, Debug)]
pub struct Suggestion {
    pub title: String,
    pub reason: String,
    pub rule: DetectionRule,
}

/// Rules derived from what the project already knows, best first.
pub fn suggest(project: &Project) -> Vec<Suggestion> {
    let mut out: Vec<Suggestion> = Vec::new();
    let all: Vec<Action> = DeploymentType::ALL
        .iter()
        .flat_map(|dt| project.actions(*dt).unwrap_or_default())
        .collect();
    let app_name = project.setting(keys::APP_NAME).trim().to_string();
    let app_version = project.setting(keys::APP_VERSION).trim().to_string();

    // 1. MSI product code
    let mut codes: Vec<(String, String)> = Vec::new();
    let remembered = project.setting(PRODUCT_CODE_KEY);
    if is_guid(remembered) {
        codes.push((normalize_guid(remembered), "remembered from the MSI import".into()));
    }
    for a in &all {
        for key in ["ProductCode", "InstalledApplication"] {
            let v = a.value(key);
            if is_guid(v) {
                codes.push((normalize_guid(v), format!("used by step \"{}\"", a.name)));
            }
        }
    }
    codes.dedup_by(|a, b| a.0 == b.0);
    for (code, why) in codes.into_iter().take(2) {
        out.push(Suggestion {
            title: "MSI product code".into(),
            reason: format!("{code} — {why}"),
            rule: DetectionRule::Msi { product_code: code, version_operator: Operator::NotConfigured, product_version: String::new() },
        });
    }

    // 2. File: target of a shortcut the package creates
    for a in all.iter().filter(|a| a.command == "New-ADTShortcut" && !a.disabled) {
        let target = a.value("TargetPath");
        if let Some(p) = to_intune_path(target) {
            if let Some((dir, file)) = p.rsplit_once('\\') {
                if !file.is_empty() && dir.contains('%') {
                    out.push(Suggestion {
                        title: "File exists".into(),
                        reason: format!("{p} — target of shortcut \"{}\"", a.name),
                        rule: DetectionRule::File {
                            path: dir.to_string(),
                            file_or_folder: file.to_string(),
                            check: FileCheck::Exists,
                            operator: Operator::NotConfigured,
                            value: String::new(),
                            check_32bit: false,
                        },
                    });
                }
            }
        }
    }

    // 3. Registry value the package writes
    for a in all.iter().filter(|a| a.command == "Set-ADTRegistryKey" && !a.disabled) {
        if let Some(key) = to_intune_registry_path(a.value("LiteralPath")) {
            let name = a.value("Name").trim().to_string();
            let value = a.value("Value").trim().trim_matches('"').trim_matches('\'').to_string();
            let (check, operator, cmp) = if !value.is_empty() && !value.contains('$') {
                if value.chars().all(|c| c.is_ascii_digit()) {
                    (RegistryCheck::Integer, Operator::Equal, value.clone())
                } else {
                    (RegistryCheck::String, Operator::Equal, value.clone())
                }
            } else {
                (RegistryCheck::Exists, Operator::NotConfigured, String::new())
            };
            out.push(Suggestion {
                title: "Registry value".into(),
                reason: format!("written by step \"{}\"", a.name),
                rule: DetectionRule::Registry { key_path: key, value_name: name, check, operator, value: cmp, check_32bit: false },
            });
        }
    }

    // 4. Generic: Add/Remove Programs entry by name and version
    if !app_name.is_empty() {
        out.push(Suggestion {
            title: "Script: Add/Remove Programs".into(),
            reason: format!(
                "\"{app_name}\"{} is listed in the Uninstall registry keys",
                if app_version.is_empty() { String::new() } else { format!(" ≥ {app_version}") }
            ),
            rule: DetectionRule::Script {
                script: uninstall_registry_script(&app_name, &app_version),
                run_as_32bit: false,
                enforce_signature_check: false,
            },
        });
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn guids() {
        assert!(is_guid("{23170F69-40C1-2702-2501-000001000000}"));
        assert!(is_guid("23170f69-40c1-2702-2501-000001000000"));
        assert!(!is_guid("{23170F69-40C1-2702-2501-00000100000}"));
        assert_eq!(normalize_guid("23170f69-40c1-2702-2501-000001000000"), "{23170F69-40C1-2702-2501-000001000000}");
    }

    #[test]
    fn paths() {
        assert_eq!(to_intune_path("$envProgramFiles\\7-Zip\\7zFM.exe").as_deref(), Some("%ProgramFiles%\\7-Zip\\7zFM.exe"));
        assert_eq!(to_intune_path("$envProgramFilesX86\\App\\a.exe").as_deref(), Some("%ProgramFiles(x86)%\\App\\a.exe"));
        assert_eq!(to_intune_path("\"$env:ProgramFiles\\X\\x.exe\"").as_deref(), Some("%ProgramFiles%\\X\\x.exe"));
        assert_eq!(to_intune_path("$InstallationFolder\\7zFM.exe"), None);
        assert_eq!(
            to_intune_registry_path("HKLM:\\SOFTWARE\\Company\\Deployments").as_deref(),
            Some("HKEY_LOCAL_MACHINE\\SOFTWARE\\Company\\Deployments")
        );
        assert_eq!(
            to_intune_registry_path("Registry::HKEY_CURRENT_USER\\Software\\X\\").as_deref(),
            Some("HKEY_CURRENT_USER\\Software\\X")
        );
        assert_eq!(to_intune_registry_path("SOFTWARE\\X"), None);
    }

    #[test]
    fn json_roundtrip_and_graph() {
        let r = DetectionRule::File {
            path: "%ProgramFiles%\\7-Zip\\".into(),
            file_or_folder: "7zFM.exe".into(),
            check: FileCheck::Version,
            operator: Operator::GreaterThanOrEqual,
            value: "25.1.0.0".into(),
            check_32bit: false,
        };
        let s = serde_json::to_string(&r).unwrap();
        assert!(s.contains("\"type\":\"file\""));
        assert_eq!(serde_json::from_str::<DetectionRule>(&s).unwrap(), r);
        let g = r.to_graph();
        assert_eq!(g["@odata.type"], "#microsoft.graph.win32LobAppFileSystemRule");
        assert_eq!(g["path"], "%ProgramFiles%\\7-Zip");
        assert_eq!(g["operationType"], "version");
        assert_eq!(g["operator"], "greaterThanOrEqual");
        assert!(r.problems().is_empty());

        let m = DetectionRule::Msi { product_code: "23170f69-40c1-2702-2501-000001000000".into(), version_operator: Operator::NotConfigured, product_version: String::new() };
        let g = m.to_graph();
        assert_eq!(g["productCode"], "{23170F69-40C1-2702-2501-000001000000}");
        assert!(g["productVersion"].is_null());

        let sc = DetectionRule::Script { script: uninstall_registry_script("7-Zip", "25.01"), run_as_32bit: false, enforce_signature_check: false };
        assert!(!sc.to_graph()["scriptContent"].as_str().unwrap().is_empty());
        let bad = DetectionRule::Registry { key_path: "SOFTWARE\\X".into(), value_name: String::new(), check: RegistryCheck::Integer, operator: Operator::NotConfigured, value: String::new(), check_32bit: false };
        assert_eq!(bad.problems().len(), 3);
    }
}
