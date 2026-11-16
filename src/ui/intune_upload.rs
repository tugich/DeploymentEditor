//! "Upload to Intune" window: settings, sign-in, progress and result.

use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::Instant;

use eframe::egui::{self, CornerRadius, Stroke, Vec2};
use egui_phosphor::regular as icon;
use serde::{Deserialize, Serialize};

use crate::app::{Cmd, DeploymentEditorApp, JobMsg, ToastKind};
use crate::detection::DetectionRule;
use crate::intune::{self, AppSpec, Auth, DeviceCodePrompt, Endpoints, ExistingApp, Graph, MinOs, Token};
use crate::project::{keys, Project, SETUP_FILE};
use crate::theme::semibold;
use crate::widgets::{field_label, it, muted, primary_button, secondary_button, segmented};

const SETTINGS_KEY: &str = "Intune_App";
const APP_ID_KEY: &str = "Intune_AppId";

/// Per-project app settings (stored as JSON in the project database).
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(default)]
pub struct IntuneAppSettings {
    pub display_name: String,
    pub description: String,
    pub publisher: String,
    pub version: String,
    pub developer: String,
    pub notes: String,
    pub install_cmd: String,
    pub uninstall_cmd: String,
    pub run_as_system: bool,
    pub restart_behavior: String,
    pub max_runtime_minutes: u32,
    pub allow_available_uninstall: bool,
    pub architectures: String,
    pub min_os: String,
    pub use_icon: bool,
}

impl Default for IntuneAppSettings {
    fn default() -> Self {
        Self {
            display_name: String::new(),
            description: "Uploaded by Deployment Editor (TUGI.CH)".into(),
            publisher: String::new(),
            version: String::new(),
            developer: String::new(),
            notes: String::new(),
            install_cmd: "Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent".into(),
            uninstall_cmd: "Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent".into(),
            run_as_system: true,
            restart_behavior: "suppress".into(),
            max_runtime_minutes: 60,
            allow_available_uninstall: true,
            architectures: "x64".into(),
            min_os: MinOs::W11_21H2.graph().into(),
            use_icon: true,
        }
    }
}

impl IntuneAppSettings {
    pub fn for_project(p: &Project) -> Self {
        let mut s: IntuneAppSettings = serde_json::from_str(p.setting(SETTINGS_KEY)).unwrap_or_default();
        // Name / version always follow the project (they change with every new version).
        let vendor = p.setting(keys::APP_VENDOR).trim();
        let name = p.setting(keys::APP_NAME).trim();
        let version = p.setting(keys::APP_VERSION).trim();
        let arch = p.setting(keys::APP_ARCH).trim();
        s.display_name = [vendor, name, version, arch].iter().filter(|x| !x.is_empty()).cloned().collect::<Vec<_>>().join(" ");
        s.publisher = if vendor.is_empty() { s.publisher } else { vendor.to_string() };
        s.version = version.to_string();
        if p.setting(SETTINGS_KEY).is_empty() {
            s.architectures = match arch.to_ascii_lowercase().as_str() {
                "x86" => "x86",
                "arm64" => "arm64",
                "x64" => "x64",
                _ => "x64,x86",
            }
            .into();
        }
        s
    }

    fn spec(&self, detection: Vec<serde_json::Value>, icon_png: Option<Vec<u8>>) -> AppSpec {
        AppSpec {
            display_name: self.display_name.clone(),
            description: self.description.clone(),
            publisher: self.publisher.clone(),
            version: self.version.clone(),
            developer: self.developer.clone(),
            notes: self.notes.clone(),
            install_cmd: self.install_cmd.clone(),
            uninstall_cmd: self.uninstall_cmd.clone(),
            run_as_system: self.run_as_system,
            restart_behavior: self.restart_behavior.clone(),
            max_runtime_minutes: self.max_runtime_minutes,
            allow_available_uninstall: self.allow_available_uninstall,
            architectures: self.architectures.clone(),
            min_os: MinOs::from_key(&self.min_os).unwrap_or(MinOs::W11_21H2),
            detection,
            icon_png: if self.use_icon { icon_png } else { None },
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Phase {
    Edit,
    Running { fraction: f32, text: String },
    Done { app_id: String, created: bool },
    Failed(String),
}

pub enum IntuneMsg {
    Progress(f32, String),
    DeviceCode(DeviceCodePrompt),
    Token(String, Token),
    Apps(Result<Vec<ExistingApp>, String>),
    Done(Result<(String, bool), String>),
}

pub struct UploadForm {
    pub use_secret: bool,
    pub tenant: String,
    pub client_id: String,
    pub public_client_id: String,
    pub secret: String,
    pub remember_secret: bool,
    pub app: IntuneAppSettings,
    /// `None` = create a new app
    pub target: Option<ExistingApp>,
    pub found: Vec<ExistingApp>,
    pub searching: bool,
    pub search_error: Option<String>,
    pub phase: Phase,
    pub device_code: Option<DeviceCodePrompt>,
    pub cancel: Arc<AtomicBool>,
    pub icon: Option<PathBuf>,
    pub started: Option<Instant>,
}

struct ConnCfg {
    use_secret: bool,
    tenant: String,
    client_id: String,
    secret: String,
}

impl ConnCfg {
    fn identity(&self) -> String {
        format!("{}|{}|{}", self.use_secret, self.tenant.trim(), self.client_id.trim())
    }
}

/// Sign in (reusing a cached token) and return a Graph client.
fn connect(cfg: &ConnCfg, cached: Option<Token>, send: &dyn Fn(IntuneMsg), cancel: &AtomicBool) -> anyhow::Result<Graph> {
    let ep = Endpoints::default();
    let auth = if cfg.use_secret {
        Auth::ClientSecret { tenant: cfg.tenant.clone(), client_id: cfg.client_id.clone(), secret: cfg.secret.clone() }
    } else {
        Auth::DeviceCode { tenant: cfg.tenant.clone(), client_id: cfg.client_id.clone() }
    };
    let token = match cached.filter(|t| t.valid_for(std::time::Duration::from_secs(600))) {
        Some(t) => t,
        None => {
            send(IntuneMsg::Progress(0.0, "Signing in to Microsoft Entra ID".into()));
            let t = if cfg.use_secret {
                intune::token_client_secret(&ep, &cfg.tenant, &cfg.client_id, &cfg.secret)?
            } else {
                let prompt = intune::device_code_start(&ep, &cfg.tenant, &cfg.client_id)?;
                send(IntuneMsg::DeviceCode(prompt.clone()));
                intune::device_code_wait(&ep, &cfg.tenant, &cfg.client_id, &prompt, &|| cancel.load(Ordering::Relaxed))?
            };
            send(IntuneMsg::Token(cfg.identity(), t.clone()));
            t
        }
    };
    Ok(Graph::new(ep, auth, token))
}

impl DeploymentEditorApp {
    fn conn_cfg(&self, f: &UploadForm) -> ConnCfg {
        ConnCfg {
            use_secret: f.use_secret,
            tenant: f.tenant.trim().to_string(),
            client_id: if f.use_secret { f.client_id.trim().to_string() } else { f.public_client_id.trim().to_string() },
            secret: f.secret.clone(),
        }
    }

    fn cached_token(&self, cfg: &ConnCfg) -> Option<Token> {
        self.intune_token.as_ref().filter(|(id, _)| *id == cfg.identity()).map(|(_, t)| t.clone())
    }

    pub fn open_upload(&mut self) {
        let Some(p) = &self.project else { return };
        let app = IntuneAppSettings::for_project(p);
        let existing_id = p.setting(APP_ID_KEY).trim().to_string();
        let icon = Some(p.folder.join("Assets").join("AppIcon.png")).filter(|i| i.is_file());
        let tenant = self.prefs.intune_tenant.clone();
        let client_id = self.prefs.intune_client_id.clone();
        let stored = if self.prefs.intune_use_secret && !tenant.is_empty() && !client_id.is_empty() {
            intune::stored_secret(&tenant, &client_id)
        } else {
            None
        };
        let target = (!existing_id.is_empty()).then(|| ExistingApp {
            id: existing_id.clone(),
            display_name: "App from the last upload".into(),
            version: String::new(),
            publisher: String::new(),
        });
        self.upload = Some(Box::new(UploadForm {
            use_secret: self.prefs.intune_use_secret,
            tenant,
            client_id,
            public_client_id: if self.prefs.intune_public_client_id.is_empty() {
                intune::DEFAULT_PUBLIC_CLIENT_ID.into()
            } else {
                self.prefs.intune_public_client_id.clone()
            },
            remember_secret: stored.is_some(),
            secret: stored.unwrap_or_default(),
            app,
            target,
            found: Vec::new(),
            searching: false,
            search_error: None,
            phase: Phase::Edit,
            device_code: None,
            cancel: Arc::new(AtomicBool::new(false)),
            icon,
            started: None,
        }));
    }

    fn spawn_intune<F>(&self, job: F)
    where
        F: FnOnce(&dyn Fn(IntuneMsg)) -> IntuneMsg + Send + 'static,
    {
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        std::thread::spawn(move || {
            let send = |m: IntuneMsg| {
                let _ = tx.send(JobMsg::Intune(m));
                ctx.request_repaint();
            };
            let last = job(&send);
            send(last);
        });
    }

    fn start_search(&mut self) {
        let Some(f) = self.upload.as_mut() else { return };
        f.searching = true;
        f.search_error = None;
        f.device_code = None;
        f.cancel.store(false, Ordering::Relaxed);
        let f = self.upload.as_ref().unwrap();
        let cfg = self.conn_cfg(f);
        let cached = self.cached_token(&cfg);
        let cancel = f.cancel.clone();
        let name = self.project.as_ref().map(|p| p.setting(keys::APP_NAME).trim().to_string()).unwrap_or_default();
        self.spawn_intune(move |send| {
            let res = connect(&cfg, cached, send, &cancel).and_then(|mut g| intune::find_apps(&mut g, &name));
            IntuneMsg::Apps(res.map_err(|e| format!("{e:#}")))
        });
    }

    fn start_upload(&mut self) {
        // Always upload what the editor shows right now.
        if !self.generate(true) {
            return;
        }
        let Some(project) = self.project.as_mut() else { return };
        let Some(f) = self.upload.as_mut() else { return };
        let Some(rule) = DetectionRule::load(project) else {
            f.phase = Phase::Failed("Set up a detection rule first (Intune tab → Detection Rule).".into());
            return;
        };
        // remember connection + app settings
        self.prefs.intune_tenant = f.tenant.trim().to_string();
        self.prefs.intune_client_id = f.client_id.trim().to_string();
        self.prefs.intune_use_secret = f.use_secret;
        self.prefs.intune_public_client_id = f.public_client_id.trim().to_string();
        if f.use_secret {
            let res = intune::store_secret(&f.tenant, &f.client_id, if f.remember_secret { Some(&f.secret) } else { None });
            if let Err(e) = res {
                self.status = format!("Could not store the client secret: {e}");
            }
        }
        let _ = project.set_setting(SETTINGS_KEY, &serde_json::to_string(&f.app).unwrap_or_default());

        let icon_png = f.icon.as_ref().and_then(|p| std::fs::read(p).ok());
        let spec = f.app.spec(vec![rule.to_graph()], icon_png);
        let existing = f.target.as_ref().map(|t| t.id.clone());
        let folder = project.folder.clone();
        f.cancel.store(false, Ordering::Relaxed);
        f.phase = Phase::Running { fraction: 0.0, text: "Starting".into() };
        f.device_code = None;
        f.started = Some(Instant::now());
        let f = self.upload.as_ref().unwrap();
        let cfg = self.conn_cfg(f);
        let cached = self.cached_token(&cfg);
        let cancel = f.cancel.clone();
        self.busy = Some("Uploading to Intune…".into());
        self.spawn_intune(move |send| {
            let work = std::env::temp_dir().join(format!("de-upload-{}", std::process::id()));
            let res = (|| -> anyhow::Result<(String, bool)> {
                let mut graph = connect(&cfg, cached, send, &cancel)?;
                let mut last = Instant::now();
                let pkg = crate::intunewin::build_encrypted(&folder, SETUP_FILE, &work, &mut |fr, t| {
                    if last.elapsed().as_millis() > 120 {
                        last = Instant::now();
                        send(IntuneMsg::Progress(0.02 + fr * 0.18, t));
                    }
                })?;
                if cancel.load(Ordering::Relaxed) {
                    anyhow::bail!("Upload cancelled.");
                }
                let id = intune::upload(
                    &mut graph,
                    &spec,
                    existing.as_deref(),
                    &pkg,
                    &mut |fr, t| send(IntuneMsg::Progress(0.2 + fr * 0.8, t)),
                    &|| cancel.load(Ordering::Relaxed),
                )?;
                Ok((id, existing.is_none()))
            })();
            let _ = std::fs::remove_dir_all(&work);
            IntuneMsg::Done(res.map_err(|e| format!("{e:#}")))
        });
    }

    pub fn handle_intune_msg(&mut self, msg: IntuneMsg) {
        match msg {
            IntuneMsg::Token(id, t) => self.intune_token = Some((id, t)),
            IntuneMsg::DeviceCode(p) => {
                if let Some(f) = self.upload.as_mut() {
                    f.device_code = Some(p);
                }
            }
            IntuneMsg::Progress(fr, text) => {
                if let Some(f) = self.upload.as_mut() {
                    if matches!(f.phase, Phase::Running { .. }) {
                        f.phase = Phase::Running { fraction: fr, text };
                    }
                }
            }
            IntuneMsg::Apps(res) => {
                if let Some(f) = self.upload.as_mut() {
                    f.searching = false;
                    f.device_code = None;
                    match res {
                        Ok(list) => {
                            if let Some(t) = f.target.as_mut() {
                                if let Some(found) = list.iter().find(|a| a.id == t.id) {
                                    *t = found.clone();
                                }
                            }
                            f.search_error = list.is_empty().then(|| "No matching Win32 app found in Intune.".to_string());
                            f.found = list;
                        }
                        Err(e) => {
                            if e.contains("returned 401") || e.contains("returned 403") {
                                self.intune_token = None;
                            }
                            f.search_error = Some(e);
                        }
                    }
                }
            }
            IntuneMsg::Done(res) => {
                self.busy = None;
                self.upload_hidden = false;
                let Some(f) = self.upload.as_mut() else { return };
                f.device_code = None;
                match res {
                    Ok((id, created)) => {
                        f.phase = Phase::Done { app_id: id.clone(), created };
                        f.target = Some(ExistingApp {
                            id: id.clone(),
                            display_name: f.app.display_name.clone(),
                            version: f.app.version.clone(),
                            publisher: f.app.publisher.clone(),
                        });
                        if let Some(p) = self.project.as_mut() {
                            let _ = p.set_setting(APP_ID_KEY, &id);
                        }
                        self.toast(ToastKind::Success, "Uploaded to Intune");
                    }
                    Err(e) => {
                        // After a permission problem the next attempt must use a new token
                        // (a token only contains the permissions granted when it was issued).
                        if e.contains("returned 401") || e.contains("returned 403") {
                            self.intune_token = None;
                        }
                        f.phase = Phase::Failed(e);
                    }
                }
            }
        }
    }

    pub(super) fn upload_window(&mut self, ctx: &egui::Context) {
        let Some(mut f) = self.upload.take() else { return };
        let p = self.palette;
        let mut open = true;
        let mut action: Option<&'static str> = None;
        let running = matches!(f.phase, Phase::Running { .. });

        egui::Window::new(it(format!("{}  Upload to Microsoft Intune", icon::CLOUD_ARROW_UP)))
            .open(&mut open)
            .pivot(egui::Align2::CENTER_CENTER)
            .default_pos(ctx.content_rect().center())
            .default_size(Vec2::new(720.0, 680.0))
            .collapsible(false)
            .resizable(true)
            .show(ctx, |ui| {
                egui::Panel::bottom("upload-footer").show_separator_line(true).show(ui, |ui| {
                    ui.add_space(8.0);
                    self.upload_footer(ui, &mut f, &mut action);
                    ui.add_space(2.0);
                });
                egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
                    ui.add_enabled_ui(!running, |ui| {
                        ui.spacing_mut().item_spacing.y = 5.0;
                        self.upload_connection(ui, &mut f, &mut action);
                        ui.add_space(10.0);
                        self.upload_target(ui, &mut f, &mut action);
                        ui.add_space(10.0);
                        self.upload_app(ui, &mut f);
                        ui.add_space(10.0);
                        self.upload_program(ui, &mut f);
                        ui.add_space(10.0);
                        self.upload_detection(ui, &mut action);
                    });
                });
            });

        if let Some(code) = &f.device_code {
            device_code_overlay(ctx, code, &p, &mut action);
        }

        if !open {
            if running {
                // Keep the job alive in the background; Upload to Intune shows it again.
                self.toast(ToastKind::Info, "The upload continues in the background");
                self.upload = Some(f);
                self.upload_hidden = true;
                return;
            }
            f.cancel.store(true, Ordering::Relaxed);
            return;
        }
        self.upload = Some(f);
        match action {
            Some("search") => self.start_search(),
            Some("upload") => self.start_upload(),
            Some("detection") => self.cmds.push(Cmd::DetectionRule),
            Some("cancel") => {
                if let Some(f) = self.upload.as_mut() {
                    f.cancel.store(true, Ordering::Relaxed);
                    f.device_code = None;
                }
            }
            _ => {}
        }
    }

    fn upload_connection(&mut self, ui: &mut egui::Ui, f: &mut UploadForm, _action: &mut Option<&'static str>) {
        let p = self.palette;
        section(ui, icon::KEY, "Connection");
        let mut mode = f.use_secret;
        segmented(ui, &mut mode, &[(true, "App registration".into()), (false, "Sign in with account".into())]);
        f.use_secret = mode;
        ui.add_space(2.0);
        ui.columns(2, |c| {
            field_label(&mut c[0], "Tenant ID", Some(if f.use_secret { "required" } else { "optional" }));
            c[0].add(egui::TextEdit::singleline(&mut f.tenant).hint_text("contoso.onmicrosoft.com or GUID").desired_width(f32::INFINITY));
            if f.use_secret {
                field_label(&mut c[1], "Client ID", None);
                c[1].add(egui::TextEdit::singleline(&mut f.client_id).hint_text("App registration (client) ID").desired_width(f32::INFINITY));
            } else {
                field_label(&mut c[1], "Client ID", Some("public client"));
                c[1].add(egui::TextEdit::singleline(&mut f.public_client_id).desired_width(f32::INFINITY));
            }
        });
        if f.use_secret {
            field_label(ui, "Client secret", None);
            ui.horizontal(|ui| {
                ui.add(egui::TextEdit::singleline(&mut f.secret).password(true).desired_width(360.0));
                ui.checkbox(&mut f.remember_secret, "Remember in Windows Credential Manager");
            });
            muted(ui, "Needs the application permission DeviceManagementApps.ReadWrite.All with admin consent.");
        } else {
            ui.label(
                it(format!(
                    "{}  You sign in with a code in the browser. Default client: Microsoft Graph Command Line Tools — your account needs an Intune role that may manage apps.",
                    icon::INFO
                ))
                .size(12.0)
                .color(p.text_muted),
            );
        }
    }

    fn upload_target(&mut self, ui: &mut egui::Ui, f: &mut UploadForm, action: &mut Option<&'static str>) {
        let p = self.palette;
        section(ui, icon::PACKAGE, "Target");
        let mut new_app = f.target.is_none();
        ui.horizontal(|ui| {
            if ui.radio(new_app, "Create a new app").clicked() {
                new_app = true;
            }
            if ui.radio(!new_app, "New version of an existing app").clicked() && new_app {
                new_app = false;
                if f.target.is_none() {
                    if let Some(first) = f.found.first() {
                        f.target = Some(first.clone());
                    } else {
                        *action = Some("search");
                    }
                }
            }
        });
        if new_app {
            f.target = None;
        }
        if !new_app || f.searching {
            ui.horizontal(|ui| {
                match &f.target {
                    Some(t) => {
                        ui.label(it(format!("{}  {}", icon::CHECK_CIRCLE, t.display_name)).font(semibold(13.0)));
                        if !t.version.is_empty() {
                            ui.label(it(format!("v{}", t.version)).size(12.0).color(p.text_muted));
                        }
                        ui.label(it(&t.id).monospace().size(11.0).color(p.text_faint));
                    }
                    None => {
                        muted(ui, "No app selected.");
                    }
                }
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if f.searching {
                        ui.add(egui::Spinner::new());
                    } else if secondary_button(ui, &format!("{}  Find in Intune", icon::MAGNIFYING_GLASS)).clicked() {
                        *action = Some("search");
                    }
                });
            });
            if !f.found.is_empty() {
                egui::Frame::new().stroke(Stroke::new(1.0, p.border)).corner_radius(CornerRadius::same(8)).inner_margin(egui::Margin::same(4)).show(ui, |ui| {
                    egui::ScrollArea::vertical().id_salt("found").max_height(130.0).show(ui, |ui| {
                        ui.set_width(ui.available_width());
                        for a in &f.found {
                            let sel = f.target.as_ref().map(|t| t.id == a.id).unwrap_or(false);
                            let label = format!("{}   {}   {}", a.display_name, a.version, a.publisher);
                            if ui.selectable_label(sel, label).clicked() {
                                f.target = Some(a.clone());
                            }
                        }
                    });
                });
            }
            if let Some(e) = &f.search_error {
                ui.add(egui::Label::new(it(format!("{}  {e}", icon::WARNING)).size(12.0).color(p.warning)).wrap());
            }
        }
    }

    fn upload_app(&mut self, ui: &mut egui::Ui, f: &mut UploadForm) {
        let p = self.palette;
        section(ui, icon::APP_WINDOW, "App information");
        ui.columns(2, |c| {
            field_label(&mut c[0], "Name", None);
            c[0].add(egui::TextEdit::singleline(&mut f.app.display_name).desired_width(f32::INFINITY));
            field_label(&mut c[1], "Publisher", None);
            c[1].add(egui::TextEdit::singleline(&mut f.app.publisher).desired_width(f32::INFINITY));
            field_label(&mut c[0], "Version", None);
            c[0].add(egui::TextEdit::singleline(&mut f.app.version).desired_width(f32::INFINITY));
            field_label(&mut c[1], "Developer", Some("optional"));
            c[1].add(egui::TextEdit::singleline(&mut f.app.developer).desired_width(f32::INFINITY));
        });
        field_label(ui, "Description", None);
        ui.add(egui::TextEdit::multiline(&mut f.app.description).desired_rows(2).desired_width(f32::INFINITY));
        ui.horizontal(|ui| {
            match &f.icon {
                Some(_) => {
                    ui.checkbox(&mut f.app.use_icon, "Use Assets\\AppIcon.png as app icon");
                }
                None => {
                    ui.label(it(format!("{}  No Assets\\AppIcon.png in the project — Intune shows a default icon.", icon::IMAGE)).size(12.0).color(p.text_muted));
                }
            }
        });
    }

    fn upload_program(&mut self, ui: &mut egui::Ui, f: &mut UploadForm) {
        section(ui, icon::TERMINAL_WINDOW, "Program & requirements");
        field_label(ui, "Install command", None);
        ui.add(egui::TextEdit::singleline(&mut f.app.install_cmd).font(egui::TextStyle::Monospace).desired_width(f32::INFINITY));
        field_label(ui, "Uninstall command", None);
        ui.add(egui::TextEdit::singleline(&mut f.app.uninstall_cmd).font(egui::TextStyle::Monospace).desired_width(f32::INFINITY));
        ui.add_space(2.0);
        egui::Grid::new("prog-grid").num_columns(4).spacing(Vec2::new(12.0, 6.0)).show(ui, |ui| {
            ui.label("Install behavior");
            egui::ComboBox::from_id_salt("runas").selected_text(if f.app.run_as_system { "System" } else { "User" }).show_ui(ui, |ui| {
                ui.selectable_value(&mut f.app.run_as_system, true, "System");
                ui.selectable_value(&mut f.app.run_as_system, false, "User");
            });
            ui.label("Restart");
            let restart = [
                ("basedOnReturnCode", "Determine by return codes"),
                ("suppress", "No specific action"),
                ("allow", "App may force a restart"),
                ("force", "Intune will force a restart"),
            ];
            let current = restart.iter().find(|r| r.0 == f.app.restart_behavior).map(|r| r.1).unwrap_or("No specific action");
            egui::ComboBox::from_id_salt("restart").selected_text(current).width(200.0).show_ui(ui, |ui| {
                for (k, l) in restart {
                    ui.selectable_value(&mut f.app.restart_behavior, k.to_string(), l);
                }
            });
            ui.end_row();
            ui.label("Architecture");
            let archs = [("x64", "64-bit (x64)"), ("x86", "32-bit (x86)"), ("arm64", "ARM64"), ("x64,x86", "x64 + x86"), ("x64,x86,arm64", "All incl. ARM64")];
            let current = archs.iter().find(|a| a.0 == f.app.architectures).map(|a| a.1).unwrap_or("x64");
            egui::ComboBox::from_id_salt("arch").selected_text(current).show_ui(ui, |ui| {
                for (k, l) in archs {
                    ui.selectable_value(&mut f.app.architectures, k.to_string(), l);
                }
            });
            ui.label("Minimum OS");
            let current = MinOs::from_key(&f.app.min_os).unwrap_or(MinOs::W11_21H2);
            egui::ComboBox::from_id_salt("minos").selected_text(current.label()).width(200.0).show_ui(ui, |ui| {
                for m in MinOs::ALL {
                    ui.selectable_value(&mut f.app.min_os, m.graph().to_string(), m.label());
                }
            });
            ui.end_row();
            ui.label("Max. run time");
            ui.add(egui::DragValue::new(&mut f.app.max_runtime_minutes).range(1..=1440).suffix(" min"));
            ui.label("");
            ui.checkbox(&mut f.app.allow_available_uninstall, "Allow uninstall in Company Portal");
            ui.end_row();
        });
    }

    fn upload_detection(&mut self, ui: &mut egui::Ui, action: &mut Option<&'static str>) {
        let p = self.palette;
        section(ui, icon::MAGNIFYING_GLASS_PLUS, "Detection rule");
        let rule = self.project.as_ref().and_then(DetectionRule::load);
        ui.horizontal(|ui| {
            match &rule {
                Some(r) if r.problems().is_empty() => {
                    ui.label(it(format!("{}  {}", icon::CHECK_CIRCLE, r.summary())).color(p.success));
                }
                Some(r) => {
                    ui.label(it(format!("{}  {}", icon::WARNING, r.problems().join(" "))).color(p.warning));
                }
                None => {
                    ui.label(it(format!("{}  No detection rule yet — required for the upload.", icon::WARNING_CIRCLE)).color(p.danger));
                }
            }
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if secondary_button(ui, if rule.is_some() { "Edit…" } else { "Set up…" }).clicked() {
                    *action = Some("detection");
                }
            });
        });
    }

    fn upload_footer(&mut self, ui: &mut egui::Ui, f: &mut UploadForm, action: &mut Option<&'static str>) {
        let p = self.palette;
        let has_rule = self.project.as_ref().and_then(DetectionRule::load).is_some();
        match &f.phase {
            Phase::Running { fraction, text } => {
                ui.add(egui::ProgressBar::new(*fraction).show_percentage().desired_width(f32::INFINITY));
                ui.horizontal(|ui| {
                    ui.add(egui::Spinner::new().size(14.0));
                    ui.label(text.as_str());
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if secondary_button(ui, "Cancel").clicked() {
                            *action = Some("cancel");
                        }
                        if let Some(s) = f.started {
                            muted(ui, &format!("{} s", s.elapsed().as_secs()));
                        }
                    });
                });
            }
            Phase::Done { app_id, created } => {
                ui.horizontal(|ui| {
                    ui.label(it(icon::CHECK_CIRCLE).size(24.0).color(p.success));
                    ui.vertical(|ui| {
                        ui.label(egui::RichText::new(if *created { "App created in Intune" } else { "New version uploaded" }).font(semibold(15.0)));
                        muted(ui, "Intune is publishing the content now — assign the app to groups in the Intune admin center.");
                    });
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if primary_button(ui, &format!("{}  Open in Intune", icon::ARROW_SQUARE_OUT)).clicked() {
                            self.cmds.push(Cmd::OpenUrl(intune::portal_url(app_id)));
                        }
                    });
                });
                if secondary_button(ui, "Upload again").clicked() {
                    f.phase = Phase::Edit;
                }
            }
            Phase::Edit | Phase::Failed(_) => {
                if let Phase::Failed(e) = &f.phase {
                    egui::Frame::new().fill(p.danger.gamma_multiply(0.1)).corner_radius(CornerRadius::same(8)).inner_margin(egui::Margin::same(10)).show(ui, |ui| {
                        ui.set_width(ui.available_width());
                        egui::ScrollArea::vertical().id_salt("err").max_height(110.0).show(ui, |ui| {
                            ui.add(egui::Label::new(it(format!("{}  {e}", icon::WARNING_CIRCLE)).color(p.danger)).wrap().selectable(true));
                        });
                    });
                    ui.add_space(6.0);
                }
                let conn_ok = if f.use_secret {
                    !f.tenant.trim().is_empty() && !f.client_id.trim().is_empty() && !f.secret.is_empty()
                } else {
                    !f.public_client_id.trim().is_empty()
                };
                let ok = conn_ok && has_rule && !f.app.display_name.trim().is_empty() && !f.searching;
                ui.horizontal(|ui| {
                    ui.add_enabled_ui(ok, |ui| {
                        let label = if f.target.is_some() { "Upload new version" } else { "Create app & upload" };
                        if primary_button(ui, &format!("{}  {label}", icon::CLOUD_ARROW_UP)).clicked() {
                            *action = Some("upload");
                        }
                    });
                    if !conn_ok {
                        muted(ui, "Enter the connection details.");
                    } else if !has_rule {
                        muted(ui, "A detection rule is required.");
                    } else {
                        muted(ui, "The deployment file is generated and packaged automatically.");
                    }
                });
            }
        }
    }
}

fn section(ui: &mut egui::Ui, ic: &str, title: &str) {
    let p = crate::widgets::palette(ui);
    ui.horizontal(|ui| {
        ui.label(it(ic).size(16.0).color(p.accent));
        ui.label(egui::RichText::new(title).font(semibold(14.5)));
    });
}

fn device_code_overlay(ctx: &egui::Context, code: &DeviceCodePrompt, p: &crate::theme::Palette, action: &mut Option<&'static str>) {
    egui::Modal::new(egui::Id::new("device-code")).show(ctx, |ui| {
        ui.set_width(460.0);
        ui.label(it(format!("{}  Sign in to Microsoft", icon::KEY)).font(semibold(18.0)));
        ui.add_space(6.0);
        ui.label("Open the sign-in page and enter this code:");
        ui.add_space(8.0);
        ui.vertical_centered(|ui| {
            ui.label(egui::RichText::new(&code.user_code).font(egui::FontId::monospace(30.0)).color(p.accent_text).strong());
        });
        ui.add_space(8.0);
        ui.horizontal(|ui| {
            if primary_button(ui, &format!("{}  Copy code & open browser", icon::ARROW_SQUARE_OUT)).clicked() {
                ui.ctx().copy_text(code.user_code.clone());
                let _ = crate::system::open_url(&code.verification_uri);
            }
            if secondary_button(ui, "Cancel").clicked() {
                *action = Some("cancel");
            }
        });
        ui.add_space(4.0);
        ui.horizontal(|ui| {
            ui.add(egui::Spinner::new().size(13.0));
            muted(ui, &format!("Waiting for sign-in at {}", code.verification_uri));
        });
    });
}
