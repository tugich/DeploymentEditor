//! Application state and all non-visual behaviour (commands).

use std::collections::HashSet;
use std::path::{Path, PathBuf};
use std::sync::mpsc::{channel, Receiver, Sender};
use std::time::Instant;

use eframe::egui;

use crate::catalog::{Catalog, CUSTOM_SCRIPT};
use crate::generator;
use crate::logs::LogSet;
use crate::paths::AppPaths;
use crate::plugins::{self, Plugin, Snippet};
use crate::project::{keys, Action, DeploymentType, Project, ProjectSettings, DATABASE_FILE};
use crate::system::{self, MsiInfo};
use crate::theme::{self, Palette};
use crate::validate::{self, Issue, Severity};

// ---------------------------------------------------------------------------- prefs

#[derive(serde::Serialize, serde::Deserialize, Clone, Debug)]
#[serde(default)]
pub struct Prefs {
    pub recent: Vec<PathBuf>,
    pub silent_mode: bool,
    pub dark_mode: bool,
    pub ribbon_collapsed: bool,
    pub show_library: bool,
    pub show_inspector: bool,
    pub last_project: Option<PathBuf>,
    pub intune_tenant: String,
    pub intune_client_id: String,
    /// true = app registration + client secret, false = device code sign-in
    pub intune_use_secret: bool,
    pub intune_public_client_id: String,
    pub remote: crate::remote::RemoteOptions,
}

impl Default for Prefs {
    fn default() -> Self {
        Self {
            recent: Vec::new(),
            silent_mode: false,
            dark_mode: false,
            ribbon_collapsed: false,
            show_library: true,
            show_inspector: true,
            last_project: None,
            intune_tenant: String::new(),
            intune_client_id: String::new(),
            intune_use_secret: true,
            intune_public_client_id: crate::intune::DEFAULT_PUBLIC_CLIENT_ID.into(),
            remote: Default::default(),
        }
    }
}

// ---------------------------------------------------------------------------- UI enums

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum RibbonTab {
    Home,
    Insert,
    Run,
    Project,
    Intune,
    Tools,
    View,
    Help,
}

impl RibbonTab {
    pub const ALL: [RibbonTab; 8] = [
        RibbonTab::Home,
        RibbonTab::Insert,
        RibbonTab::Run,
        RibbonTab::Project,
        RibbonTab::Intune,
        RibbonTab::Tools,
        RibbonTab::View,
        RibbonTab::Help,
    ];
    pub fn label(&self) -> &'static str {
        match self {
            RibbonTab::Home => "Home",
            RibbonTab::Insert => "Insert",
            RibbonTab::Run => "Run",
            RibbonTab::Project => "Project",
            RibbonTab::Intune => "Intune",
            RibbonTab::Tools => "Tools",
            RibbonTab::View => "View",
            RibbonTab::Help => "Help",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum BackstagePage {
    Home,
    New,
    Open,
    Info,
    About,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum InspectorTab {
    Options,
    Parameters,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum FolderKind {
    Project,
    Files,
    SupportFiles,
    Logs,
    Snippets,
    Plugins,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Editor {
    Ise,
    NotepadPlusPlus,
    VsCode,
}

#[derive(Clone, Debug, PartialEq)]
pub enum NewKind {
    Blank,
    Exe(String),
    Msi(String),
}

pub struct NewProjectForm {
    pub kind: NewKind,
    pub folder: String,
    pub name: String,
    pub msi: Option<MsiInfo>,
    pub msi_loading: bool,
    pub confirm_overwrite: bool,
    pub error: Option<String>,
}

pub struct WinGetForm {
    pub query: String,
    pub searching: bool,
    pub results: Vec<crate::winget::Hit>,
    pub selected: Option<usize>,
    pub id: String,
    pub version: String,
    pub manifest: Option<crate::winget::Manifest>,
    pub loading_manifest: bool,
    pub installer: Option<usize>,
    pub arch: String,
    pub folder: String,
    pub confirm_overwrite: bool,
    pub creating: bool,
    pub error: Option<String>,
}

impl Default for WinGetForm {
    fn default() -> Self {
        Self {
            query: String::new(),
            searching: false,
            results: Vec::new(),
            selected: None,
            id: String::new(),
            version: String::new(),
            manifest: None,
            loading_manifest: false,
            installer: None,
            arch: "x64".into(),
            folder: String::new(),
            confirm_overwrite: false,
            creating: false,
            error: None,
        }
    }
}

pub struct DetectionForm {
    pub rule: Option<crate::detection::DetectionRule>,
    pub suggestions: Vec<crate::detection::Suggestion>,
}

pub enum Dialog {
    /// Problems list; `then` is the command to run on "Continue anyway".
    Problems { then: Option<Box<Cmd>> },
    Detection(DetectionForm),
    WinGet(WinGetForm),
    ProjectSettings(ProjectSettings),
    NewProject(NewProjectForm),
    ConfirmRemove(i64, String),
    Message { title: String, body: String, error: bool },
    ExecutablesList(String),
    Plugins,
    ConfirmUnloadForPlugin(usize),
    ExportPdf(crate::report::ReportOptions),
    Remote(RemoteForm),
}

pub struct RemoteForm {
    pub opts: crate::remote::RemoteOptions,
    pub error: Option<String>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ToastKind {
    Info,
    Success,
    Error,
}

pub struct Toast {
    pub text: String,
    pub kind: ToastKind,
    pub created: Instant,
}

pub struct Preview {
    pub full: bool,
    pub text: String,
}

pub struct LogViewer {
    pub logs: LogSet,
    pub filter: String,
    pub show_info: bool,
    pub show_warning: bool,
    pub show_error: bool,
    pub selected: Option<usize>,
}

/// One step in the undo history: the project state *before* a change.
pub struct UndoEntry {
    pub label: String,
    pub snapshot: crate::project::Snapshot,
    pub dt: DeploymentType,
    pub selected: Option<i64>,
}

const UNDO_LIMIT: usize = 100;

pub enum JobMsg {
    Status(String),
    Done { ok: bool, title: String, body: String },
    Msi(Result<MsiInfo, String>),
    WinGetSearch(Result<Vec<crate::winget::Hit>, String>),
    WinGetManifest(Result<crate::winget::Manifest, String>),
    WinGetCreated(Result<PathBuf, String>),
    PsCheck(u64, Result<Vec<Issue>, String>),
    Intune(crate::ui::intune_upload::IntuneMsg),
}

/// Everything the UI can ask for; executed after the frame is drawn.
#[derive(Clone, Debug)]
pub enum Cmd {
    Backstage(Option<BackstagePage>),
    NewProject(NewKind),
    WinGetImport,
    BrowseOpenProject,
    OpenProject(PathBuf),
    OpenDemoProject,
    CloseProject,
    Reload,
    Save,
    ProjectSettings,
    SetDeploymentType(DeploymentType),
    Select(Option<i64>),
    AddCommand(String),
    AddCommandWith(String, String, Vec<(String, String)>),
    AddCustomScript,
    AddSnippet(usize),
    DropFiles(Vec<PathBuf>),
    RemoveSelected,
    ConfirmedRemove(i64),
    Duplicate,
    Undo,
    Redo,
    DetectionRule,
    ShowProblems,
    IntuneUpload,
    RemoteRun,
    RemoteRunNow(crate::remote::RemoteOptions),
    ExportPdf,
    ExportPdfNow(crate::report::ReportOptions),
    GoTo(DeploymentType, Option<i64>),
    ToggleDisable(Option<i64>),
    Move(i32),
    MoveTo(i64, usize),
    Generate,
    Preview,
    Run(DeploymentType),
    Sandbox,
    Intune,
    PowerShellHere,
    OpenFolder(FolderKind),
    OpenWith(Editor),
    ExecutablesList,
    ShowPlugins,
    RunPlugin(usize),
    HelpConsole,
    OpenUrl(String),
    CommandHelp,
    LogViewer,
    ToggleDark,
    ExpandAll(bool),
    ForgetRecent(PathBuf),
    Quit,
}

// ---------------------------------------------------------------------------- app

pub struct DeploymentEditorApp {
    pub paths: AppPaths,
    pub catalog: Catalog,
    pub catalog_error: Option<String>,
    pub project: Option<Project>,
    pub project_name: String,
    pub dt: DeploymentType,
    pub actions: Vec<Action>,
    pub counts: [usize; 3],
    pub selected: Option<i64>,
    pub draft: Option<Action>,
    pub cond_ui: crate::ui::condition_builder::CondUi,
    /// Renderer and adapter, shown on the About page.
    pub graphics: String,
    pub dirty: bool,
    pub expanded: HashSet<i64>,
    pub ribbon_tab: RibbonTab,
    pub backstage: Option<BackstagePage>,
    pub lib_search: String,
    pub lib_selected: Option<String>,
    pub quick_search: String,
    pub inspector_tab: InspectorTab,
    pub param_search: String,
    pub only_set_params: bool,
    pub prefs: Prefs,
    pub dialog: Option<Dialog>,
    pub toasts: Vec<Toast>,
    pub status: String,
    pub snippets: Vec<Snippet>,
    pub plugins: Vec<Plugin>,
    pub tx: Sender<JobMsg>,
    pub rx: Receiver<JobMsg>,
    pub busy: Option<String>,
    pub preview: Option<Preview>,
    pub script_editor_open: bool,
    pub log_viewer: Option<LogViewer>,
    pub last_generated: Option<String>,
    pub palette: Palette,
    pub cmds: Vec<Cmd>,
    pub ctx: egui::Context,
    pub window_title: String,
    pub scroll_to_selected: bool,
    pub pending_insert_at: Option<usize>,
    pub quick_popup_hovered: bool,
    pub undo_stack: Vec<UndoEntry>,
    pub redo_stack: Vec<UndoEntry>,
    /// Results of the static checks (recomputed on every change).
    pub issues: Vec<Issue>,
    /// Syntax errors found by the PowerShell parser (Windows).
    pub ps_issues: Vec<Issue>,
    pub ps_check_due: Option<Instant>,
    pub ps_check_gen: u64,
    pub ps_check_running: bool,
    pub ps_check_error: Option<String>,
    /// Set by "Continue anyway" for exactly one gated command.
    pub gate_override: bool,
    pub upload: Option<Box<crate::ui::intune_upload::UploadForm>>,
    pub upload_hidden: bool,
    /// Cached Microsoft Graph token: (connection identity, token)
    pub intune_token: Option<(String, crate::intune::Token)>,
}

const PREFS_KEY: &str = "deployment-editor-prefs";

impl DeploymentEditorApp {
    pub fn new(cc: &eframe::CreationContext<'_>) -> Self {
        let graphics = graphics_info(cc);
        let paths = AppPaths::detect();
        let (catalog, catalog_error) = match Catalog::load(&paths.psadt_database()) {
            Ok(c) => (c, None),
            Err(e) => (Catalog::default(), Some(format!("{e:#}"))),
        };
        let prefs: Prefs = cc
            .storage
            .and_then(|s| eframe::get_value(s, PREFS_KEY))
            .unwrap_or_default();
        theme::install_fonts(&cc.egui_ctx);
        let palette = Palette::for_mode(prefs.dark_mode);
        theme::apply_style(&cc.egui_ctx, &palette);
        cc.egui_ctx.data_mut(|d| d.insert_temp(egui::Id::new("palette"), palette));

        let (tx, rx) = channel();
        let mut app = Self {
            snippets: plugins::load_snippets(&paths.snippets()),
            plugins: plugins::load_plugins(&paths.plugins()),
            paths,
            catalog,
            catalog_error,
            project: None,
            project_name: String::new(),
            dt: DeploymentType::Installation,
            actions: Vec::new(),
            counts: [0; 3],
            selected: None,
            draft: None,
            cond_ui: Default::default(),
            graphics,
            dirty: false,
            expanded: HashSet::new(),
            ribbon_tab: RibbonTab::Home,
            backstage: Some(BackstagePage::Home),
            lib_search: String::new(),
            lib_selected: None,
            quick_search: String::new(),
            inspector_tab: InspectorTab::Options,
            param_search: String::new(),
            only_set_params: false,
            prefs,
            dialog: None,
            toasts: Vec::new(),
            status: "Ready".into(),
            tx,
            rx,
            busy: None,
            preview: None,
            script_editor_open: false,
            log_viewer: None,
            last_generated: None,
            palette,
            cmds: Vec::new(),
            ctx: cc.egui_ctx.clone(),
            window_title: String::new(),
            scroll_to_selected: false,
            pending_insert_at: None,
            quick_popup_hovered: false,
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            issues: Vec::new(),
            ps_issues: Vec::new(),
            ps_check_due: None,
            ps_check_gen: 0,
            ps_check_running: false,
            ps_check_error: None,
            gate_override: false,
            upload: None,
            upload_hidden: false,
            intune_token: None,
        };

        // Command line: DeploymentEditor.exe <project.db | project folder>
        if let Some(arg) = std::env::args().nth(1) {
            let p = PathBuf::from(arg);
            let db = if p.is_dir() { p.join(DATABASE_FILE) } else { p };
            app.open_project(&db);
        }
        if let Some(err) = app.catalog_error.clone() {
            app.dialog = Some(Dialog::Message {
                title: "PSADT command database missing".into(),
                body: format!(
                    "{err}\n\nPlace the Rust executable in the DeploymentEditor folder (next to 'Databases', 'Templates' and 'ThirdParty') or set DEPLOYMENT_EDITOR_HOME."
                ),
                error: true,
            });
        }
        app
    }

    pub fn save_prefs(&self, storage: &mut dyn eframe::Storage) {
        eframe::set_value(storage, PREFS_KEY, &self.prefs);
    }

    // ------------------------------------------------------------------ feedback

    pub fn toast(&mut self, kind: ToastKind, text: impl Into<String>) {
        let text = text.into();
        self.status = text.clone();
        self.toasts.push(Toast { text, kind, created: Instant::now() });
        if self.toasts.len() > 4 {
            self.toasts.remove(0);
        }
    }

    pub fn error(&mut self, title: &str, err: impl std::fmt::Display) {
        self.status = format!("{title}: {err}");
        self.dialog = Some(Dialog::Message { title: title.into(), body: err.to_string(), error: true });
    }

    pub fn has_project(&self) -> bool {
        self.project.is_some()
    }

    pub fn selected_action(&self) -> Option<&Action> {
        let id = self.selected?;
        self.actions.iter().find(|a| a.id == id)
    }

    // ------------------------------------------------------------------ project lifecycle

    pub fn open_project(&mut self, db: &Path) {
        self.commit_draft();
        match Project::open(db) {
            Ok(p) => {
                self.project_name = p.display_name();
                let db = p.database.clone();
                self.project = Some(p);
                self.clear_history();
                self.selected = None;
                self.draft = None;
                self.expanded.clear();
                self.dt = DeploymentType::Installation;
                self.preview = None;
                self.last_generated = None;
                self.backstage = None;
                self.push_recent(db.clone());
                self.prefs.last_project = Some(db);
                self.refresh();
                let n = self.counts.iter().sum::<usize>();
                self.toast(ToastKind::Success, format!("Opened \"{}\" — {n} actions", self.project_name));
            }
            Err(e) => self.error("Can't open project", format!("{e:#}")),
        }
    }

    fn push_recent(&mut self, db: PathBuf) {
        self.prefs.recent.retain(|p| p != &db);
        self.prefs.recent.insert(0, db);
        self.prefs.recent.truncate(10);
    }

    pub fn close_project(&mut self) {
        self.commit_draft();
        self.project = None;
        self.upload = None;
        self.clear_history();
        self.actions.clear();
        self.selected = None;
        self.draft = None;
        self.counts = [0; 3];
        self.preview = None;
        self.script_editor_open = false;
        self.backstage = Some(BackstagePage::Home);
        self.status = "Project closed".into();
    }

    /// Re-read the action list for the current deployment type.
    pub fn refresh(&mut self) {
        let Some(p) = &self.project else { return };
        match p.actions(self.dt) {
            Ok(list) => self.actions = list,
            Err(e) => {
                let msg = format!("{e:#}");
                self.error("Database error", msg);
                return;
            }
        }
        for (i, dt) in DeploymentType::ALL.iter().enumerate() {
            self.counts[i] = p.action_count(*dt);
        }
        if let Some(id) = self.selected {
            if !self.actions.iter().any(|a| a.id == id) {
                self.selected = None;
                self.draft = None;
                self.dirty = false;
            } else if !self.dirty {
                self.draft = self.actions.iter().find(|a| a.id == id).cloned();
            }
        }
        self.revalidate();
    }

    /// Re-run the static checks and schedule the PowerShell syntax check.
    pub fn revalidate(&mut self) {
        let Some(p) = &self.project else {
            self.issues.clear();
            self.ps_issues.clear();
            return;
        };
        self.issues = validate::validate(p, &self.catalog);
        if cfg!(windows) {
            self.ps_check_due = Some(Instant::now() + std::time::Duration::from_millis(1200));
        }
    }

    /// Static issues plus PowerShell parser errors.
    pub fn all_issues(&self) -> Vec<Issue> {
        let mut v = self.issues.clone();
        v.extend(self.ps_issues.iter().cloned());
        v.sort_by_key(|i| (i.severity, i.dt.map(|d| d as u8).unwrap_or(0), i.step.unwrap_or(0)));
        v
    }

    pub fn issues_for(&self, id: i64) -> Vec<Issue> {
        self.issues.iter().chain(self.ps_issues.iter()).filter(|i| i.action_id == Some(id)).cloned().collect()
    }

    pub fn error_count(&self) -> usize {
        validate::count(&self.issues, Severity::Error) + validate::count(&self.ps_issues, Severity::Error)
    }

    fn start_ps_check(&mut self) {
        let Some(p) = &self.project else { return };
        let mut items: Vec<(i64, String)> = Vec::new();
        let mut meta: Vec<(i64, DeploymentType, usize, String)> = Vec::new();
        for dt in DeploymentType::ALL {
            for (idx, a) in p.actions(dt).unwrap_or_default().into_iter().enumerate() {
                if a.disabled {
                    continue;
                }
                let mut code = generator::build_script(std::slice::from_ref(&a), &self.catalog);
                code = code.replace("\r\n", "\n");
                items.push((a.id, code));
                meta.push((a.id, dt, idx, a.name.clone()));
            }
        }
        self.ps_check_gen += 1;
        self.ps_check_running = true;
        let generation = self.ps_check_gen;
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        std::thread::spawn(move || {
            let res = system::powershell_syntax_check(&items).map(|errors| {
                errors
                    .into_iter()
                    .filter_map(|(id, line, message)| {
                        let (_, dt, idx, name) = meta.iter().find(|m| m.0 == id)?.clone();
                        Some(Issue {
                            severity: Severity::Error,
                            dt: Some(dt),
                            action_id: Some(id),
                            step: Some(idx + 1),
                            action_name: name,
                            message: format!("PowerShell: {message} (line {line} of the generated code)"),
                        })
                    })
                    .collect()
            });
            let _ = tx.send(JobMsg::PsCheck(generation, res.map_err(|e| format!("{e:#}"))));
            ctx.request_repaint();
        });
    }

    /// Returns true (and shows the problems) if errors should stop `cmd`.
    fn gated(&mut self, cmd: &Cmd) -> bool {
        if self.gate_override {
            self.gate_override = false;
            return false;
        }
        self.commit_draft();
        self.revalidate();
        if self.error_count() == 0 {
            return false;
        }
        self.dialog = Some(Dialog::Problems { then: Some(Box::new(cmd.clone())) });
        true
    }

    pub fn create_project(&mut self, form: &NewProjectForm) -> anyhow::Result<()> {
        let folder = PathBuf::from(form.folder.trim());
        if form.folder.trim().is_empty() {
            anyhow::bail!("Please choose a destination folder.");
        }
        self.commit_draft();
        let mut project = Project::create(&self.paths, &folder)?;
        let name = if form.name.trim().is_empty() { "New Project".to_string() } else { form.name.trim().to_string() };
        project.set_setting(keys::PROJECT_NAME, &name)?;
        project.set_setting(keys::APP_ARCH, "x64")?;
        project.set_setting(keys::APP_LANGUAGE, "EN")?;

        match &form.kind {
            NewKind::Blank => {
                project.set_setting(keys::APP_VERSION, "1.0")?;
                project.set_setting(keys::APP_AUTHOR, "Deployment Editor")?;
            }
            NewKind::Exe(installer) => {
                let src = PathBuf::from(installer);
                let file = copy_installer(&src, &project.files_folder())?;
                project.set_setting(keys::APP_VERSION, "1.0")?;
                project.set_setting(keys::APP_VENDOR, "Your vendor")?;
                project.set_setting(keys::APP_NAME, "App name")?;
                project.set_setting(keys::APP_AUTHOR, "Executable Import by Deployment Editor")?;
                project.add_action(
                    DeploymentType::Installation,
                    "Start-ADTProcess",
                    "Start Installer",
                    "Start the installer",
                    &[("FilePath", &file)],
                )?;
            }
            NewKind::Msi(installer) => {
                let src = PathBuf::from(installer);
                let file = copy_installer(&src, &project.files_folder())?;
                let info = form.msi.clone().unwrap_or_default();
                project.set_setting(keys::APP_VERSION, &info.product_version)?;
                project.set_setting(keys::APP_VENDOR, &info.manufacturer)?;
                project.set_setting(keys::APP_NAME, &info.product_name)?;
                project.set_setting(keys::APP_AUTHOR, "MSI Import by Deployment Editor")?;
                if crate::detection::is_guid(&info.product_code) {
                    project.set_setting(crate::detection::PRODUCT_CODE_KEY, &crate::detection::normalize_guid(&info.product_code))?;
                }
                project.add_action(
                    DeploymentType::Installation,
                    "Start-ADTMsiProcess",
                    "Run MSI installer",
                    "Start the installer",
                    &[("FilePath", &file), ("Action", "Install")],
                )?;
                if !info.product_code.is_empty() {
                    project.add_action(
                        DeploymentType::Uninstall,
                        "Start-ADTMsiProcess",
                        "Run MSI uninstall",
                        "Start the uninstall",
                        &[("ProductCode", &info.product_code), ("Action", "Uninstall")],
                    )?;
                } else {
                    project.add_action(
                        DeploymentType::Uninstall,
                        "Start-ADTMsiProcess",
                        "Run MSI uninstall",
                        "Start the uninstall",
                        &[("FilePath", &file), ("Action", "Uninstall")],
                    )?;
                }
            }
        }
        let db = project.database.clone();
        drop(project);
        self.open_project(&db);
        Ok(())
    }

    // ------------------------------------------------------------------ selection & draft

    pub fn select(&mut self, id: Option<i64>) {
        if self.selected == id {
            return;
        }
        self.commit_draft();
        self.selected = id;
        self.draft = id.and_then(|id| self.actions.iter().find(|a| a.id == id).cloned());
        self.dirty = false;
        self.script_editor_open = false;
    }

    /// Persist pending edits of the selected action.
    pub fn commit_draft(&mut self) {
        if !self.dirty {
            return;
        }
        self.dirty = false;
        let label = match self.draft.as_ref() {
            Some(d) => format!("Edit \"{}\"", d.name),
            None => return,
        };
        self.checkpoint(label);
        let (Some(project), Some(draft)) = (self.project.as_mut(), self.draft.as_ref()) else { return };
        let mut a = draft.clone();
        // Make sure every catalog parameter has a row (same as the original Save).
        if !a.is_custom_script() {
            for p in self.catalog.parameters(&a.command) {
                if !a.values.iter().any(|(n, _)| n == &p.name) {
                    let v = if p.is_checkbox() { "0" } else { "" };
                    a.values.push((p.name.clone(), v.into()));
                }
            }
        } else {
            let script = plugins::to_crlf(a.value("Script"));
            a.set_value("Script", script);
        }
        match project.save_action(&a) {
            Ok(()) => {
                self.status = format!("Saved \"{}\"", a.name);
                self.draft = Some(a);
                self.refresh();
            }
            Err(e) => self.error("Save failed", format!("{e:#}")),
        }
    }

    /// "Set-ADTRegistryKey" -> "Set Registry Key"
    pub fn new_action_name(command: &str) -> String {
        let (verb, noun) = command.split_once('-').unwrap_or((command, ""));
        let noun = noun.strip_prefix("ADT").unwrap_or(noun);
        let mut words = String::new();
        let chars: Vec<char> = noun.chars().collect();
        for (i, c) in chars.iter().enumerate() {
            let boundary = i > 0
                && c.is_uppercase()
                && (chars[i - 1].is_lowercase() || chars.get(i + 1).map(|n| n.is_lowercase()).unwrap_or(false));
            if boundary {
                words.push(' ');
            }
            words.push(*c);
        }
        format!("{verb} {words}").trim().to_string()
    }

    fn after_add(&mut self, id: i64, tab: InspectorTab) {
        self.refresh();
        self.select(Some(id));
        self.expanded.insert(id);
        self.inspector_tab = tab;
        self.scroll_to_selected = true;
        self.prefs.show_inspector = true;
    }

    pub fn add_command(&mut self, command: &str, name: Option<String>, values: Vec<(String, String)>) {
        self.commit_draft();
        let name = name.unwrap_or_else(|| Self::new_action_name(command));
        self.checkpoint(format!("Add \"{name}\""));
        let Some(project) = self.project.as_mut() else { return };
        let refs: Vec<(&str, &str)> = values.iter().map(|(a, b)| (a.as_str(), b.as_str())).collect();
        match project.add_action(self.dt, command, &name, "", &refs) {
            Ok(id) => {
                self.after_add(id, InspectorTab::Parameters);
                self.toast(ToastKind::Info, format!("Added {command}"));
            }
            Err(e) => self.error("Can't add action", format!("{e:#}")),
        }
    }

    pub fn add_custom_script(&mut self, name: &str, script: &str) {
        self.commit_draft();
        self.checkpoint(format!("Add \"{name}\""));
        let Some(project) = self.project.as_mut() else { return };
        match project.add_action(self.dt, CUSTOM_SCRIPT, name, "", &[("Script", script)]) {
            Ok(id) => {
                self.after_add(id, InspectorTab::Parameters);
                self.toast(ToastKind::Info, format!("Added \"{name}\""));
            }
            Err(e) => self.error("Can't add script", format!("{e:#}")),
        }
    }

    pub fn add_snippet(&mut self, idx: usize) {
        let Some(snippet) = self.snippets.get(idx).cloned() else { return };
        match plugins::read_text_file(&snippet.path) {
            Ok(mut text) => {
                if !text.ends_with("\r\n") {
                    text.push_str("\r\n");
                }
                self.add_custom_script(&format!("PowerShell script ({})", snippet.file_name), &text)
            }
            Err(e) => self.error("Can't read snippet", e),
        }
    }

    /// Files dropped onto the window become actions (same rules as the original
    /// drop handler, plus MSI / PS1 support).
    pub fn drop_files(&mut self, files: Vec<PathBuf>) {
        if self.project.is_none() {
            // Dropping a project database opens it.
            if let Some(db) = files.iter().find(|f| f.extension().map(|e| e == "db").unwrap_or(false)) {
                let db = db.clone();
                self.open_project(&db);
            }
            return;
        }
        let folder = self.project.as_ref().unwrap().folder.clone();
        let files_dir = self.project.as_ref().unwrap().files_folder();
        for f in files {
            let ext = f.extension().map(|e| e.to_string_lossy().to_lowercase()).unwrap_or_default();
            if ext == "db" {
                self.open_project(&f);
                return;
            }
            let rel = f
                .strip_prefix(&files_dir)
                .or_else(|_| f.strip_prefix(&folder))
                .map(|p| p.to_string_lossy().to_string())
                .unwrap_or_else(|_| f.to_string_lossy().to_string());
            let file_name = f.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
            match ext.as_str() {
                "lnk" => self.add_command(
                    "Remove-ADTFile",
                    Some(format!("Remove Shortcut ({rel})")),
                    vec![("Path".into(), rel)],
                ),
                "msi" => self.add_command(
                    "Start-ADTMsiProcess",
                    Some(format!("Install {file_name}")),
                    vec![("FilePath".into(), rel), ("Action".into(), "Install".into())],
                ),
                "ps1" => match plugins::read_text_file(&f) {
                    Ok(t) => self.add_custom_script(&format!("PowerShell script ({file_name})"), &t),
                    Err(e) => self.error("Can't read script", e),
                },
                _ => self.add_command(
                    "Start-ADTProcess",
                    Some(format!("Run {file_name}")),
                    vec![("FilePath".into(), rel)],
                ),
            }
        }
    }

    // ------------------------------------------------------------------ editing

    pub fn remove_action(&mut self, id: i64) {
        let name = self.actions.iter().find(|a| a.id == id).map(|a| a.name.clone()).unwrap_or_default();
        self.checkpoint(format!("Remove \"{name}\""));
        let Some(project) = self.project.as_mut() else { return };
        if let Err(e) = project.remove_action(id) {
            return self.error("Can't remove action", format!("{e:#}"));
        }
        if self.selected == Some(id) {
            self.selected = None;
            self.draft = None;
            self.dirty = false;
        }
        self.refresh();
        self.toast(ToastKind::Info, "Action removed — Ctrl+Z to undo");
    }

    pub fn duplicate_selected(&mut self) {
        self.commit_draft();
        if self.selected.is_some() {
            self.checkpoint("Duplicate step");
        }
        let (Some(project), Some(id)) = (self.project.as_mut(), self.selected) else { return };
        match project.duplicate_action(id) {
            Ok(Some(new_id)) => self.after_add(new_id, self.inspector_tab),
            Ok(None) => {}
            Err(e) => self.error("Can't duplicate", format!("{e:#}")),
        }
    }

    pub fn toggle_disable(&mut self, id: Option<i64>) {
        let Some(id) = id.or(self.selected) else { return };
        if Some(id) == self.selected {
            if let Some(d) = self.draft.as_mut() {
                d.disabled = !d.disabled;
                self.dirty = true;
                self.commit_draft();
                return;
            }
        }
        let disabled = self.actions.iter().find(|a| a.id == id).map(|a| a.disabled).unwrap_or(false);
        self.checkpoint(if disabled { "Enable step" } else { "Disable step" });
        let Some(project) = self.project.as_mut() else { return };
        if let Err(e) = project.set_disabled(id, !disabled) {
            return self.error("Database error", format!("{e:#}"));
        }
        self.refresh();
    }

    pub fn move_selected(&mut self, delta: i32) {
        self.commit_draft();
        if self.selected.is_some() {
            self.checkpoint("Move step");
        }
        let (Some(project), Some(id)) = (self.project.as_mut(), self.selected) else { return };
        match project.move_action(id, self.dt, delta) {
            Ok(_) => {
                self.refresh();
                self.scroll_to_selected = true;
            }
            Err(e) => self.error("Can't move action", format!("{e:#}")),
        }
    }

    pub fn move_to(&mut self, id: i64, index: usize, record: bool) {
        self.commit_draft();
        if record {
            self.checkpoint("Move step");
        }
        let Some(project) = self.project.as_mut() else { return };
        match project.move_action_to(id, self.dt, index) {
            Ok(_) => self.refresh(),
            Err(e) => self.error("Can't move action", format!("{e:#}")),
        }
    }

    // ------------------------------------------------------------------ undo / redo

    /// Record the current project state before a change.
    pub fn checkpoint(&mut self, label: impl Into<String>) {
        let Some(project) = &self.project else { return };
        match project.snapshot() {
            Ok(snapshot) => {
                self.undo_stack.push(UndoEntry { label: label.into(), snapshot, dt: self.dt, selected: self.selected });
                if self.undo_stack.len() > UNDO_LIMIT {
                    self.undo_stack.remove(0);
                }
                self.redo_stack.clear();
            }
            Err(e) => self.status = format!("Undo history unavailable: {e:#}"),
        }
    }

    fn clear_history(&mut self) {
        self.undo_stack.clear();
        self.redo_stack.clear();
    }

    pub fn undo_label(&self) -> Option<&str> {
        self.undo_stack.last().map(|e| e.label.as_str())
    }

    pub fn redo_label(&self) -> Option<&str> {
        self.redo_stack.last().map(|e| e.label.as_str())
    }

    pub fn undo(&mut self) {
        self.commit_draft();
        self.step_history(true);
    }

    pub fn redo(&mut self) {
        self.commit_draft();
        self.step_history(false);
    }

    fn step_history(&mut self, undo: bool) {
        let entry = if undo { self.undo_stack.pop() } else { self.redo_stack.pop() };
        let Some(entry) = entry else { return };
        let Some(project) = self.project.as_mut() else { return };
        let current = match project.snapshot() {
            Ok(s) => s,
            Err(e) => return self.error("Undo failed", format!("{e:#}")),
        };
        if let Err(e) = project.restore(&entry.snapshot) {
            return self.error("Undo failed", format!("{e:#}"));
        }
        let counter = UndoEntry { label: entry.label.clone(), snapshot: current, dt: self.dt, selected: self.selected };
        if undo {
            self.redo_stack.push(counter);
        } else {
            self.undo_stack.push(counter);
        }
        self.project_name = project.display_name();
        self.dt = entry.dt;
        self.selected = None;
        self.draft = None;
        self.dirty = false;
        self.refresh();
        if let Some(id) = entry.selected {
            if self.actions.iter().any(|a| a.id == id) {
                self.select(Some(id));
                self.scroll_to_selected = true;
            }
        }
        let verb = if undo { "Undone" } else { "Redone" };
        self.toast(ToastKind::Info, format!("{verb}: {}", entry.label));
    }

    pub fn save_detection(&mut self, rule: Option<&crate::detection::DetectionRule>) {
        self.checkpoint(if rule.is_some() { "Detection rule" } else { "Remove detection rule" });
        let Some(p) = self.project.as_mut() else { return };
        match crate::detection::DetectionRule::save(p, rule) {
            Ok(()) => {
                self.refresh();
                self.toast(ToastKind::Success, if rule.is_some() { "Detection rule saved" } else { "Detection rule removed" });
            }
            Err(e) => self.error("Can't save the detection rule", format!("{e:#}")),
        }
    }

    pub fn set_deployment_type(&mut self, dt: DeploymentType) {
        if self.dt == dt {
            return;
        }
        self.commit_draft();
        self.dt = dt;
        self.selected = None;
        self.draft = None;
        self.refresh();
    }

    // ------------------------------------------------------------------ generate & run

    pub fn generate(&mut self, quiet: bool) -> bool {
        self.commit_draft();
        let Some(project) = &self.project else { return false };
        match generator::generate_deployment_file(project, &self.catalog, &self.paths.template_script()) {
            Ok(()) => {
                let now = chrono::Local::now().format("%H:%M:%S").to_string();
                self.last_generated = Some(now);
                if !quiet {
                    self.toast(
                        ToastKind::Success,
                        format!("Deployment file generated: {}", project.deployment_file().display()),
                    );
                }
                true
            }
            Err(e) => {
                self.error("Can't generate the deployment file", format!("{e:#}"));
                false
            }
        }
    }

    pub fn build_preview(&mut self, full: bool) {
        self.commit_draft();
        let Some(project) = &self.project else { return };
        let text = if full {
            let tpl = if project.deployment_template_file().is_file() {
                project.deployment_template_file()
            } else {
                self.paths.template_script()
            };
            let raw = std::fs::read(&tpl).map(|b| String::from_utf8_lossy(&b).to_string()).unwrap_or_default();
            let today = chrono::Local::now().format("%Y-%m-%d").to_string();
            generator::render_template(&raw, project, &self.catalog, &today)
                .map(|s| s.trim_start_matches('\u{feff}').to_string())
                .unwrap_or_else(|e| format!("# Error: {e:#}"))
        } else {
            let actions = project.actions_for_generation(self.dt).unwrap_or_default();
            generator::build_script(&actions, &self.catalog)
        };
        self.preview = Some(Preview { full, text: text.replace("\r\n", "\n") });
    }

    pub fn run(&mut self, dt: DeploymentType) {
        if !self.generate(true) {
            return;
        }
        let Some(project) = &self.project else { return };
        let exe = project.folder.join("Invoke-AppDeployToolkit.exe");
        let mut args = format!("-DeploymentType '{}'", dt.as_psadt());
        if self.prefs.silent_mode {
            args.push_str(" -DeployMode Silent");
        }
        let folder = project.folder.clone();
        match system::run_elevated(&exe, &args, &folder) {
            Ok(()) => self.toast(ToastKind::Success, format!("Started deployment › {}", dt.as_psadt())),
            Err(e) => self.error("Can't start the deployment", e),
        }
    }

    pub fn sandbox(&mut self) {
        if !self.generate(true) {
            return;
        }
        let Some(project) = &self.project else { return };
        match system::start_sandbox(&self.paths.sandbox_template(), &project.folder) {
            Ok(()) => self.toast(ToastKind::Success, "Starting Windows Sandbox…"),
            Err(e) => self.error("Windows Sandbox", e),
        }
    }

    /// Export `Invoke-AppDeployToolkit.intunewin` (native packager, no IntuneWinAppUtil needed).
    pub fn intune(&mut self) {
        if !self.generate(true) {
            return;
        }
        let Some(project) = &self.project else { return };
        let Some(dest) = rfd::FileDialog::new().set_title("Intune package destination").pick_folder() else {
            return;
        };
        let folder = project.folder.clone();
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        self.busy = Some("Creating Intune package…".into());
        std::thread::spawn(move || {
            let mut last = Instant::now();
            let res = crate::intunewin::create_package_file(&folder, crate::project::SETUP_FILE, &dest, &mut |f, text| {
                if last.elapsed().as_millis() > 150 {
                    last = Instant::now();
                    let _ = tx.send(JobMsg::Status(format!("{text} — {:.0} %", f * 100.0)));
                    ctx.request_repaint();
                }
            });
            let msg = match res {
                Ok(path) => JobMsg::Done {
                    ok: true,
                    title: "Intune package created".into(),
                    body: format!("Intune package created: {}", path.display()),
                },
                Err(e) => JobMsg::Done { ok: false, title: "Intune package creation failed".into(), body: format!("{e:#}") },
            };
            let _ = tx.send(msg);
            ctx.request_repaint();
        });
    }

    /// Copy the package to remote computers and start the deployment there
    /// (generated PowerShell controller in a visible console window).
    pub fn remote_run(&mut self, opts: crate::remote::RemoteOptions) {
        if !self.generate(true) {
            return;
        }
        let Some(project) = &self.project else { return };
        self.prefs.remote = opts.clone();
        let folder = project.folder.clone();
        let title = format!(
            "{} {}",
            project.setting(crate::project::keys::APP_NAME).trim(),
            project.setting(crate::project::keys::APP_VERSION).trim()
        )
        .trim()
        .to_string();
        let work_root = std::env::temp_dir().join("DeploymentEditor-Remote");
        let logs_root = std::env::var_os("LOCALAPPDATA")
            .map(PathBuf::from)
            .unwrap_or_else(std::env::temp_dir)
            .join("DeploymentEditor")
            .join("RemoteLogs");
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        self.busy = Some("Preparing remote deployment…".into());
        std::thread::spawn(move || {
            let res = crate::remote::prepare(&folder, &title, &opts, &work_root, &logs_root)
                .and_then(|p| system::run_remote_controller(&p.controller).map(|_| p.computers));
            let msg = match res {
                Ok(computers) => JobMsg::Done {
                    ok: true,
                    title: String::new(),
                    body: format!("Remote deployment started on {} — see the PowerShell window", computers.join(", ")),
                },
                Err(e) => JobMsg::Done { ok: false, title: "Run on Remote Machine".into(), body: format!("{e:#}") },
            };
            let _ = tx.send(msg);
            ctx.request_repaint();
        });
    }

    /// Write the installation documentation of the open project as PDF.
    pub fn export_pdf(&mut self, opts: crate::report::ReportOptions) {
        self.commit_draft();
        self.revalidate();
        let Some(project) = &self.project else { return };
        let name = format!(
            "{} {} - Installation procedure.pdf",
            project.setting(crate::project::keys::APP_NAME).trim(),
            project.setting(crate::project::keys::APP_VERSION).trim()
        )
        .trim()
        .replace(['\\', '/', ':', '*', '?', '"', '<', '>', '|'], "_");
        let name = if name.starts_with('-') || name.starts_with(' ') { "Installation procedure.pdf".to_string() } else { name };
        let Some(path) = rfd::FileDialog::new()
            .set_title("Export documentation as PDF")
            .set_directory(&project.folder)
            .set_file_name(name.trim_start())
            .add_filter("PDF document", &["pdf"])
            .save_file()
        else {
            return;
        };
        let issues = self.all_issues();
        match crate::report::build(project, &self.catalog, &issues, opts).and_then(|b| Ok(std::fs::write(&path, b)?)) {
            Ok(()) => {
                self.toast(ToastKind::Success, "PDF documentation exported");
                let _ = system::open_path(&path);
            }
            Err(e) => self.error("Export PDF", e),
        }
    }

    pub fn read_msi_async(&mut self, path: String) {
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        let scripts = self.paths.scripts();
        std::thread::spawn(move || {
            let res = system::read_msi(&scripts, Path::new(&path)).map_err(|e| format!("{e:#}"));
            let _ = tx.send(JobMsg::Msi(res));
            ctx.request_repaint();
        });
    }

    pub fn winget_search_async(&mut self, query: String) {
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        std::thread::spawn(move || {
            let res = crate::winget::search(&query).map_err(|e| format!("{e:#}"));
            let _ = tx.send(JobMsg::WinGetSearch(res));
            ctx.request_repaint();
        });
    }

    pub fn winget_manifest_async(&mut self, id: String, version: String) {
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        std::thread::spawn(move || {
            let res = crate::winget::fetch_manifest(id.trim(), version.trim()).map_err(|e| format!("{e:#}"));
            let _ = tx.send(JobMsg::WinGetManifest(res));
            ctx.request_repaint();
        });
    }

    pub fn winget_create_async(&mut self, folder: PathBuf, manifest: crate::winget::Manifest, installer: usize) {
        self.commit_draft();
        let tx = self.tx.clone();
        let ctx = self.ctx.clone();
        let paths = self.paths.clone();
        self.busy = Some("Downloading installer from WinGet…".into());
        std::thread::spawn(move || {
            let res = create_winget_project(&paths, &folder, &manifest, installer).map_err(|e| format!("{e:#}"));
            let _ = tx.send(JobMsg::WinGetCreated(res));
            ctx.request_repaint();
        });
    }

    pub fn poll_jobs(&mut self) {
        if let Some(due) = self.ps_check_due {
            if Instant::now() >= due && !self.ps_check_running {
                self.ps_check_due = None;
                self.start_ps_check();
            } else {
                self.ctx.request_repaint_after(std::time::Duration::from_millis(300));
            }
        }
        while let Ok(msg) = self.rx.try_recv() {
            match msg {
                JobMsg::Status(s) => {
                    self.busy = Some(s.clone());
                    self.status = s;
                }
                JobMsg::Done { ok, title, body } => {
                    self.busy = None;
                    if ok {
                        self.toast(ToastKind::Success, body);
                    } else {
                        self.dialog = Some(Dialog::Message { title, body, error: true });
                    }
                }
                JobMsg::WinGetSearch(res) => {
                    if let Some(Dialog::WinGet(f)) = self.dialog.as_mut() {
                        f.searching = false;
                        match res {
                            Ok(hits) => {
                                f.error = if hits.is_empty() { Some("No package found.".into()) } else { None };
                                f.results = hits;
                                f.selected = None;
                            }
                            Err(e) => f.error = Some(e),
                        }
                    }
                }
                JobMsg::WinGetManifest(res) => {
                    if let Some(Dialog::WinGet(f)) = self.dialog.as_mut() {
                        f.loading_manifest = false;
                        match res {
                            Ok(m) => {
                                f.installer = m.best_for(&f.arch);
                                f.error = None;
                                f.manifest = Some(m);
                            }
                            Err(e) => {
                                f.manifest = None;
                                f.error = Some(e);
                            }
                        }
                    }
                }
                JobMsg::WinGetCreated(res) => {
                    self.busy = None;
                    match res {
                        Ok(db) => {
                            self.dialog = None;
                            self.open_project(&db);
                            self.generate(true);
                            self.toast(ToastKind::Success, "WinGet import complete — installer downloaded and sequence created");
                        }
                        Err(e) => {
                            if let Some(Dialog::WinGet(f)) = self.dialog.as_mut() {
                                f.creating = false;
                                f.error = Some(e);
                            } else {
                                self.error("WinGet import failed", e);
                            }
                        }
                    }
                }
                JobMsg::Intune(m) => self.handle_intune_msg(m),
                JobMsg::PsCheck(generation, res) => {
                    self.ps_check_running = false;
                    if generation == self.ps_check_gen {
                        match res {
                            Ok(list) => {
                                self.ps_issues = list;
                                self.ps_check_error = None;
                            }
                            Err(e) => {
                                self.status = format!("PowerShell syntax check unavailable: {e}");
                                self.ps_check_error = Some(e);
                            }
                        }
                    }
                }
                JobMsg::Msi(res) => {
                    if let Some(Dialog::NewProject(form)) = self.dialog.as_mut() {
                        form.msi_loading = false;
                        match res {
                            Ok(info) => {
                                if form.name.trim().is_empty() || form.name == "New Import Project" {
                                    form.name = format!("{} {}", info.product_name, info.product_version)
                                        .trim()
                                        .to_string();
                                }
                                form.msi = Some(info);
                                form.error = None;
                            }
                            Err(e) => form.error = Some(e),
                        }
                    }
                }
            }
        }
    }

    // ------------------------------------------------------------------ misc

    pub fn open_folder(&mut self, kind: FolderKind) {
        let path = match kind {
            FolderKind::Logs => crate::logs::default_log_folder(),
            FolderKind::Snippets => self.paths.snippets(),
            FolderKind::Plugins => self.paths.plugins(),
            _ => {
                let Some(p) = &self.project else { return };
                match kind {
                    FolderKind::Project => p.folder.clone(),
                    FolderKind::Files => p.files_folder(),
                    FolderKind::SupportFiles => p.support_files_folder(),
                    _ => unreachable!(),
                }
            }
        };
        if let Err(e) = system::open_path(&path) {
            self.error("Can't open folder", e);
        }
    }

    pub fn open_with(&mut self, editor: Editor) {
        let Some(p) = &self.project else { return };
        let file = p.deployment_file();
        let res = match editor {
            Editor::Ise => system::run_elevated(
                Path::new("powershell_ise.exe"),
                &format!("\"{}\"", file.display()),
                &p.folder,
            ),
            Editor::NotepadPlusPlus => system::run_program("notepad++.exe", &file),
            Editor::VsCode => system::run_program("code", &file),
        };
        if let Err(e) = res {
            self.error("Can't start the editor", e);
        }
    }

    pub fn run_plugin(&mut self, idx: usize, confirmed: bool) {
        let Some(plugin) = self.plugins.get(idx).cloned() else { return };
        let Some(project) = &self.project else {
            return self.error("Plugins", "Open a project first.");
        };
        if plugin.unload_project && !confirmed {
            self.dialog = Some(Dialog::ConfirmUnloadForPlugin(idx));
            return;
        }
        let folder = project.folder.clone();
        if plugin.unload_project {
            self.close_project();
        }
        match system::run_plugin(&plugin.script_path(), &folder) {
            Ok(()) => self.toast(ToastKind::Info, format!("Started plugin \"{}\"", plugin.name)),
            Err(e) => self.error("Can't run plugin", e),
        }
    }

    pub fn set_dark(&mut self, dark: bool) {
        self.prefs.dark_mode = dark;
        self.palette = Palette::for_mode(dark);
        theme::apply_style(&self.ctx, &self.palette);
        let pal = self.palette;
        self.ctx.data_mut(|d| d.insert_temp(egui::Id::new("palette"), pal));
    }

    pub fn execute(&mut self, cmd: Cmd) {
        match cmd {
            Cmd::Backstage(page) => {
                self.commit_draft();
                self.backstage = page;
            }
            Cmd::NewProject(kind) => {
                self.dialog = Some(Dialog::NewProject(NewProjectForm {
                    name: match &kind {
                        NewKind::Blank => "New Project".into(),
                        _ => "New Import Project".into(),
                    },
                    kind,
                    folder: String::new(),
                    msi: None,
                    msi_loading: false,
                    confirm_overwrite: false,
                    error: None,
                }))
            }
            Cmd::WinGetImport => self.dialog = Some(Dialog::WinGet(WinGetForm::default())),
            Cmd::BrowseOpenProject => {
                if let Some(f) = rfd::FileDialog::new()
                    .set_title("Open project database")
                    .add_filter("Project Database", &["db"])
                    .add_filter("All files", &["*"])
                    .pick_file()
                {
                    self.open_project(&f);
                }
            }
            Cmd::OpenProject(p) => self.open_project(&p),
            Cmd::OpenDemoProject => {
                let db = self.paths.demo_project().join(DATABASE_FILE);
                self.open_project(&db);
            }
            Cmd::CloseProject => self.close_project(),
            Cmd::Reload => {
                self.dirty = false;
                if let Some(p) = self.project.as_mut() {
                    let _ = p.reload_settings();
                    self.project_name = p.display_name();
                }
                self.draft = None;
                let sel = self.selected.take();
                self.refresh();
                self.select(sel);
                self.toast(ToastKind::Info, "Project reloaded");
            }
            Cmd::Save => {
                let was_dirty = self.dirty;
                self.commit_draft();
                if self.has_project() {
                    self.toast(ToastKind::Success, if was_dirty { "Changes saved" } else { "Everything is saved" });
                }
            }
            Cmd::ProjectSettings => {
                if let Some(p) = &self.project {
                    self.dialog = Some(Dialog::ProjectSettings(p.settings()));
                }
            }
            Cmd::SetDeploymentType(dt) => self.set_deployment_type(dt),
            Cmd::Select(id) => self.select(id),
            Cmd::AddCommand(c) => {
                self.add_command(&c, None, Vec::new());
                if let (Some(target), Some(id)) = (self.pending_insert_at.take(), self.selected) {
                    self.move_to(id, target, false);
                    self.scroll_to_selected = true;
                }
            }
            Cmd::AddCommandWith(c, name, values) => self.add_command(&c, Some(name), values),
            Cmd::AddCustomScript => self.add_custom_script("My custom PowerShell script", ""),
            Cmd::AddSnippet(i) => self.add_snippet(i),
            Cmd::DropFiles(f) => self.drop_files(f),
            Cmd::RemoveSelected => {
                if let Some(a) = self.selected_action() {
                    self.dialog = Some(Dialog::ConfirmRemove(a.id, a.name.clone()));
                }
            }
            Cmd::ConfirmedRemove(id) => self.remove_action(id),
            Cmd::Duplicate => self.duplicate_selected(),
            Cmd::DetectionRule => {
                self.commit_draft();
                if let Some(p) = &self.project {
                    self.dialog = Some(Dialog::Detection(DetectionForm {
                        rule: crate::detection::DetectionRule::load(p),
                        suggestions: crate::detection::suggest(p),
                    }));
                }
            }
            Cmd::IntuneUpload => {
                let running = self.upload.as_ref().map(|u| matches!(u.phase, crate::ui::intune_upload::Phase::Running { .. })).unwrap_or(false);
                if running {
                    self.upload_hidden = false;
                } else if !self.gated(&Cmd::IntuneUpload) {
                    self.open_upload();
                }
            }
            Cmd::ExportPdf => {
                self.commit_draft();
                if self.project.is_some() {
                    self.backstage = None;
                    self.dialog = Some(Dialog::ExportPdf(crate::report::ReportOptions::default()));
                }
            }
            Cmd::ExportPdfNow(opts) => self.export_pdf(opts),
            Cmd::RemoteRun => {
                if self.project.is_some() && !self.gated(&Cmd::RemoteRun) {
                    let mut opts = self.prefs.remote.clone();
                    opts.deployment = Some(self.dt);
                    self.dialog = Some(Dialog::Remote(RemoteForm { opts, error: None }));
                }
            }
            Cmd::RemoteRunNow(opts) => self.remote_run(opts),
            Cmd::ShowProblems => {
                self.commit_draft();
                self.revalidate();
                if cfg!(windows) {
                    self.ps_check_due = Some(Instant::now());
                }
                self.dialog = Some(Dialog::Problems { then: None });
            }
            Cmd::GoTo(dt, id) => {
                self.backstage = None;
                self.set_deployment_type(dt);
                if let Some(id) = id {
                    self.select(Some(id));
                    self.expanded.insert(id);
                    self.scroll_to_selected = true;
                    self.prefs.show_inspector = true;
                }
            }
            Cmd::Undo => self.undo(),
            Cmd::Redo => self.redo(),
            Cmd::ToggleDisable(id) => self.toggle_disable(id),
            Cmd::Move(d) => self.move_selected(d),
            Cmd::MoveTo(id, idx) => self.move_to(id, idx, true),
            Cmd::Generate => {
                if !self.gated(&Cmd::Generate) {
                    self.generate(false);
                }
            }
            Cmd::Preview => self.build_preview(false),
            Cmd::Run(dt) => {
                if !self.gated(&Cmd::Run(dt)) {
                    self.run(dt);
                }
            }
            Cmd::Sandbox => {
                if !self.gated(&Cmd::Sandbox) {
                    self.sandbox();
                }
            }
            Cmd::Intune => {
                if !self.gated(&Cmd::Intune) {
                    self.intune();
                }
            }
            Cmd::PowerShellHere => {
                if let Some(p) = &self.project {
                    let f = p.folder.clone();
                    if let Err(e) = system::powershell_here(&f) {
                        self.error("PowerShell", e);
                    }
                }
            }
            Cmd::OpenFolder(k) => self.open_folder(k),
            Cmd::OpenWith(e) => self.open_with(e),
            Cmd::ExecutablesList => {
                if let Some(dir) = rfd::FileDialog::new()
                    .set_title("Path to the installation folder")
                    .set_directory("C:\\Program Files")
                    .pick_folder()
                {
                    self.dialog = Some(Dialog::ExecutablesList(system::executables_list(&dir)));
                }
            }
            Cmd::ShowPlugins => {
                self.plugins = plugins::load_plugins(&self.paths.plugins());
                self.dialog = Some(Dialog::Plugins);
            }
            Cmd::RunPlugin(i) => self.run_plugin(i, false),
            Cmd::HelpConsole => {
                if let Err(e) = system::help_console(&self.paths.root) {
                    self.error("PSADT Help", e);
                }
            }
            Cmd::OpenUrl(u) => {
                if let Err(e) = system::open_url(&u) {
                    self.error("Can't open browser", e);
                }
            }
            Cmd::CommandHelp => {
                let cmd = self
                    .selected_action()
                    .map(|a| a.command.clone())
                    .filter(|c| !c.starts_with('#'))
                    .or_else(|| self.lib_selected.clone());
                let url = match cmd {
                    Some(c) => Catalog::help_url(&c),
                    None => "https://psappdeploytoolkit.com/docs/reference".into(),
                };
                let _ = system::open_url(&url);
            }
            Cmd::LogViewer => {
                if self.log_viewer.is_none() {
                    self.log_viewer = Some(LogViewer {
                        logs: LogSet::default(),
                        filter: String::new(),
                        show_info: true,
                        show_warning: true,
                        show_error: true,
                        selected: None,
                    });
                }
            }
            Cmd::ToggleDark => {
                let d = !self.prefs.dark_mode;
                self.set_dark(d);
            }
            Cmd::ExpandAll(open) => {
                if open {
                    self.expanded = self.actions.iter().map(|a| a.id).collect();
                } else {
                    self.expanded.clear();
                }
            }
            Cmd::ForgetRecent(p) => self.prefs.recent.retain(|r| r != &p),
            Cmd::Quit => {
                self.commit_draft();
                self.ctx.send_viewport_cmd(egui::ViewportCommand::Close);
            }
        }
    }

    pub fn run_queued(&mut self) {
        let cmds = std::mem::take(&mut self.cmds);
        for c in cmds {
            self.execute(c);
        }
    }

    pub fn window_title(&self) -> String {
        match &self.project {
            Some(_) => format!("{} — Deployment Editor", self.project_name),
            None => "Deployment Editor — TUGI.CH".into(),
        }
    }
}

/// Create a project from a WinGet manifest: copies the framework, downloads the
/// installer to `Files\Installer.<ext>` and adds the install (and MSI uninstall) step.
pub fn create_winget_project(
    paths: &AppPaths,
    folder: &Path,
    manifest: &crate::winget::Manifest,
    installer: usize,
) -> anyhow::Result<PathBuf> {
    let ins = manifest
        .installers
        .get(installer)
        .ok_or_else(|| anyhow::anyhow!("No installer selected."))?;
    let Some(ext) = ins.local_extension() else {
        anyhow::bail!(
            "The installer type '{}' is not supported for automated sequence generation (only EXE and MSI).",
            ins.installer_type
        );
    };
    let mut project = Project::create(paths, folder)?;
    let file = format!("Installer.{ext}");
    std::fs::create_dir_all(project.files_folder())?;
    crate::winget::download_installer(&ins.url, &project.files_folder().join(&file))?;

    let (vendor, name) = manifest.vendor_and_name();
    project.set_setting(keys::PROJECT_NAME, &format!("{name} {}", manifest.version))?;
    project.set_setting(keys::APP_NAME, &name)?;
    project.set_setting(keys::APP_VERSION, &manifest.version)?;
    project.set_setting(keys::APP_VENDOR, &vendor)?;
    project.set_setting(keys::APP_ARCH, &ins.architecture)?;
    project.set_setting(keys::APP_LANGUAGE, "EN")?;
    project.set_setting(keys::APP_AUTHOR, "WinGet Import by Deployment Editor")?;

    if ext == "msi" {
        project.add_action(
            DeploymentType::Installation,
            "Start-ADTMsiProcess",
            "Start MSI Installer",
            "Running the setup installer",
            &[("FilePath", &file), ("Action", "Install")],
        )?;
        let code = ins.product_code.trim();
        if crate::detection::is_guid(code) {
            project.set_setting(crate::detection::PRODUCT_CODE_KEY, &crate::detection::normalize_guid(code))?;
        }
        if code.starts_with('{') {
            project.add_action(
                DeploymentType::Uninstall,
                "Start-ADTMsiProcess",
                "Run MSI uninstall",
                "Uninstall by product code",
                &[("ProductCode", code), ("Action", "Uninstall")],
            )?;
        }
    } else {
        let silent = ins.silent_switch();
        project.add_action(
            DeploymentType::Installation,
            "Start-ADTProcess",
            "Start EXE Installer",
            "Running the setup installer",
            &[("FilePath", &file), ("ArgumentList", &silent)],
        )?;
    }
    Ok(project.database.clone())
}

fn copy_installer(src: &Path, files_dir: &Path) -> anyhow::Result<String> {
    if !src.is_file() {
        anyhow::bail!("Installer not found: {}", src.display());
    }
    std::fs::create_dir_all(files_dir)?;
    let name = src.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    std::fs::copy(src, files_dir.join(&name))?;
    Ok(name)
}

#[cfg(test)]
mod tests {
    use super::DeploymentEditorApp as A;
    #[test]
    fn friendly_names() {
        assert_eq!(A::new_action_name("Set-ADTRegistryKey"), "Set Registry Key");
        assert_eq!(A::new_action_name("Show-ADTInstallationWelcome"), "Show Installation Welcome");
        assert_eq!(A::new_action_name("Start-ADTMsiProcess"), "Start Msi Process");
        assert_eq!(A::new_action_name("ConvertTo-ADTNTAccountOrSID"), "ConvertTo NT Account Or SID");
    }
}

/// "DirectX 12 · NVIDIA …" / "OpenGL · 4.6 Mesa …" for the About page.
fn graphics_info(cc: &eframe::CreationContext<'_>) -> String {
    if let Some(rs) = &cc.wgpu_render_state {
        let info = rs.adapter.get_info();
        let backend = match info.backend {
            eframe::wgpu::Backend::Dx12 => "DirectX 12",
            eframe::wgpu::Backend::Vulkan => "Vulkan",
            eframe::wgpu::Backend::Metal => "Metal",
            eframe::wgpu::Backend::Gl => "OpenGL (wgpu)",
            _ => "wgpu",
        };
        let software = if info.device_type == eframe::wgpu::DeviceType::Cpu { " (software rendering)" } else { "" };
        return format!("{backend} · {}{software}", info.name);
    }
    if let Some(gl) = &cc.gl {
        use eframe::glow::HasContext;
        let (version, renderer) = unsafe { (gl.get_parameter_string(eframe::glow::VERSION), gl.get_parameter_string(eframe::glow::RENDERER)) };
        return format!("OpenGL {version} · {renderer}");
    }
    "unknown".into()
}
