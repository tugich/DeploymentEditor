//! The ribbon body (groups of large and small buttons per tab).

use eframe::egui::{self, Stroke};
use egui_phosphor::regular as icon;

use crate::app::{Cmd, DeploymentEditorApp, Editor, FolderKind, NewKind, RibbonTab};
use crate::project::DeploymentType;
use crate::widgets::{it, ribbon_group, ribbon_large, ribbon_small, ribbon_stack, ribbon_toggle};

/// Quick-insert templates on the Insert tab: (icon, label, command, default name, values)
pub const QUICK_INSERTS: &[(&str, &str, &str, &str, &[(&str, &str)])] = &[
    (icon::APP_WINDOW, "Run EXE", "Start-ADTProcess", "Run installer", &[("FilePath", "setup.exe"), ("ArgumentList", "/S")]),
    (icon::DISC, "Install MSI", "Start-ADTMsiProcess", "Install MSI package", &[("Action", "Install"), ("FilePath", "setup.msi")]),
    (icon::TRASH_SIMPLE, "Uninstall App", "Uninstall-ADTApplication", "Uninstall application", &[("Name", "")]),
    (icon::LINK, "Shortcut", "New-ADTShortcut", "Create shortcut", &[("LiteralPath", "$envCommonDesktop\\App.lnk"), ("TargetPath", "")]),
    (icon::TREE_STRUCTURE, "Registry Value", "Set-ADTRegistryKey", "Set registry value", &[("LiteralPath", "HKEY_LOCAL_MACHINE\\SOFTWARE\\"), ("Name", ""), ("Value", "")]),
    (icon::FILES, "Copy File", "Copy-ADTFile", "Copy files", &[("Path", "$adtSession.DirFiles\\"), ("Destination", "")]),
    (icon::FOLDER_PLUS, "New Folder", "New-ADTFolder", "Create folder", &[("LiteralPath", "")]),
    (icon::FILE_X, "Remove File", "Remove-ADTFile", "Remove file", &[("LiteralPath", "")]),
];

pub const QUICK_DIALOGS: &[(&str, &str, &str, &str, &[(&str, &str)])] = &[
    (icon::HAND_WAVING, "Welcome", "Show-ADTInstallationWelcome", "Show welcome / close apps", &[]),
    (icon::SPINNER_GAP, "Progress", "Show-ADTInstallationProgress", "Show progress", &[]),
    (icon::CHAT_TEXT, "Prompt", "Show-ADTInstallationPrompt", "Show prompt", &[("Message", ""), ("ButtonRightText", "OK")]),
    (icon::ARROW_CLOCKWISE, "Restart Prompt", "Show-ADTInstallationRestartPrompt", "Ask for restart", &[]),
];

impl DeploymentEditorApp {
    pub(super) fn ribbon(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        egui::Panel::top("ribbon")
            .exact_size(104.0)
            .frame(
                egui::Frame::new()
                    .fill(p.ribbon_bg)
                    .inner_margin(egui::Margin { left: 10, right: 10, top: 4, bottom: 4 })
                    .stroke(Stroke::new(1.0, p.border)),
            )
            .show(ui, |ui| {
                egui::ScrollArea::horizontal().scroll_bar_visibility(egui::scroll_area::ScrollBarVisibility::AlwaysHidden).show(ui, |ui| {
                    ui.horizontal_top(|ui| {
                        ui.spacing_mut().item_spacing.x = 4.0;
                        match self.ribbon_tab {
                            RibbonTab::Home => self.ribbon_home(ui),
                            RibbonTab::Insert => self.ribbon_insert(ui),
                            RibbonTab::Run => self.ribbon_run(ui),
                            RibbonTab::Project => self.ribbon_project(ui),
                            RibbonTab::Intune => self.ribbon_intune(ui),
                            RibbonTab::Tools => self.ribbon_tools(ui),
                            RibbonTab::View => self.ribbon_view(ui),
                            RibbonTab::Help => self.ribbon_help(ui),
                        }
                    });
                });
            });
    }

    fn ribbon_home(&mut self, ui: &mut egui::Ui) {
        let has_sel = self.selected.is_some();
        let lib_cmd = self.lib_selected.clone();
        let disabled_sel = self.selected_action().map(|a| a.disabled).unwrap_or(false);

        ribbon_group(ui, "Project", |ui| {
            if ribbon_large(ui, icon::FLOPPY_DISK, "Save", true, false)
                .on_hover_text("Save changes of the selected action (Ctrl+S)")
                .clicked()
            {
                self.cmds.push(Cmd::Save);
            }
            ribbon_stack(ui, |ui| {
                if ribbon_small(ui, icon::GEAR_SIX, "Settings", true).on_hover_text("App name, version, vendor…").clicked() {
                    self.cmds.push(Cmd::ProjectSettings);
                }
                if ribbon_small(ui, icon::ARROWS_CLOCKWISE, "Reload", true).clicked() {
                    self.cmds.push(Cmd::Reload);
                }
                if ribbon_small(ui, icon::FOLDER_OPEN, "Open Folder", true).clicked() {
                    self.cmds.push(Cmd::OpenFolder(FolderKind::Project));
                }
            });
        });

        ribbon_group(ui, "Actions", |ui| {
            let add_label = "Add Action";
            let tip = match &lib_cmd {
                Some(c) => format!("Add {c} (double-click in the library works too)"),
                None => "Select a command in the library first".into(),
            };
            if ribbon_large(ui, icon::PLUS_CIRCLE, add_label, lib_cmd.is_some(), false).on_hover_text(tip).clicked() {
                if let Some(c) = lib_cmd.clone() {
                    self.cmds.push(Cmd::AddCommand(c));
                }
            }
            if ribbon_large(ui, icon::CODE, "Custom Script", true, false)
                .on_hover_text("Add an inline PowerShell script")
                .clicked()
            {
                self.cmds.push(Cmd::AddCustomScript);
            }
            let snippets_resp = ribbon_large(ui, icon::BRACKETS_CURLY, &format!("Snippet {}", icon::CARET_DOWN), !self.snippets.is_empty(), false)
                .on_hover_text("Insert a script from the Snippets folder");
            egui::Popup::menu(&snippets_resp).show(|ui| {
                ui.set_min_width(240.0);
                for (i, s) in self.snippets.iter().enumerate() {
                    if ui.button(it(format!("{}  {}", icon::FILE_CODE, s.file_name))).clicked() {
                        self.cmds.push(Cmd::AddSnippet(i));
                    }
                }
                ui.separator();
                if ui.button(it(format!("{}  Open Snippets folder", icon::FOLDER_OPEN))).clicked() {
                    self.cmds.push(Cmd::OpenFolder(FolderKind::Snippets));
                }
            });
        });

        ribbon_group(ui, "Arrange", |ui| {
            ribbon_stack(ui, |ui| {
                if ribbon_small(ui, icon::ARROW_UP, "Move Up", has_sel).on_hover_text("Alt+Up").clicked() {
                    self.cmds.push(Cmd::Move(-1));
                }
                if ribbon_small(ui, icon::ARROW_DOWN, "Move Down", has_sel).on_hover_text("Alt+Down").clicked() {
                    self.cmds.push(Cmd::Move(1));
                }
                if ribbon_small(ui, icon::COPY, "Duplicate", has_sel).on_hover_text("Ctrl+D").clicked() {
                    self.cmds.push(Cmd::Duplicate);
                }
            });
            ribbon_stack(ui, |ui| {
                let (ic, label) = if disabled_sel { (icon::EYE, "Enable") } else { (icon::EYE_SLASH, "Disable") };
                if ribbon_small(ui, ic, label, has_sel).on_hover_text("Ctrl+E").clicked() {
                    self.cmds.push(Cmd::ToggleDisable(None));
                }
                if ribbon_small(ui, icon::TRASH, "Remove", has_sel).on_hover_text("Delete").clicked() {
                    self.cmds.push(Cmd::RemoveSelected);
                }
                if ribbon_small(ui, icon::QUESTION, "Command Help", has_sel || lib_cmd.is_some()).on_hover_text("F1").clicked() {
                    self.cmds.push(Cmd::CommandHelp);
                }
            });
        });

        ribbon_group(ui, "Build", |ui| {
            if ribbon_large(ui, icon::FILE_ARROW_DOWN, "Generate", true, false)
                .on_hover_text("Write Invoke-AppDeployToolkit.ps1 (Ctrl+G)")
                .clicked()
            {
                self.cmds.push(Cmd::Generate);
            }
            if ribbon_large(ui, icon::EYE, "Preview", true, false)
                .on_hover_text("Preview the generated script (Ctrl+P)")
                .clicked()
            {
                self.cmds.push(Cmd::Preview);
            }
            let errors = self.error_count();
            let label = if errors > 0 { format!("Check ({errors})") } else { "Check".to_string() };
            if ribbon_large(ui, icon::LIST_CHECKS, &label, true, false)
                .on_hover_text("Check the package for problems (F7)")
                .clicked()
            {
                self.cmds.push(Cmd::ShowProblems);
            }
        });

        ribbon_group(ui, "Test", |ui| {
            if ribbon_large(ui, icon::PLAY, "Install", true, true)
                .on_hover_text("Generate and run the installation elevated (F5)")
                .clicked()
            {
                self.cmds.push(Cmd::Run(DeploymentType::Installation));
            }
            ribbon_stack(ui, |ui| {
                ribbon_toggle(ui, &mut self.prefs.silent_mode, "Silent mode")
                    .on_hover_text("Run PSADT with -DeployMode Silent (no user interaction)");
                if ribbon_small(ui, icon::CUBE, "Windows Sandbox", true).clicked() {
                    self.cmds.push(Cmd::Sandbox);
                }
                if ribbon_small(ui, icon::TERMINAL_WINDOW, "PowerShell Here", true).clicked() {
                    self.cmds.push(Cmd::PowerShellHere);
                }
            });
        });
    }

    fn ribbon_insert(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "Script", |ui| {
            if ribbon_large(ui, icon::CODE, "Custom Script", true, false).clicked() {
                self.cmds.push(Cmd::AddCustomScript);
            }
            let r = ribbon_large(ui, icon::BRACKETS_CURLY, &format!("Snippet {}", icon::CARET_DOWN), !self.snippets.is_empty(), false);
            egui::Popup::menu(&r).show(|ui| {
                ui.set_min_width(240.0);
                for (i, s) in self.snippets.iter().enumerate() {
                    if ui.button(it(format!("{}  {}", icon::FILE_CODE, s.file_name))).clicked() {
                        self.cmds.push(Cmd::AddSnippet(i));
                    }
                }
            });
        });
        self.quick_group(ui, "Common Actions", QUICK_INSERTS);
        self.quick_group(ui, "User Interface", QUICK_DIALOGS);
    }

    fn quick_group(&mut self, ui: &mut egui::Ui, caption: &str, items: &[(&str, &str, &str, &str, &[(&str, &str)])]) {
        ribbon_group(ui, caption, |ui| {
            for (ic, label, command, name, values) in items {
                let exists = self.catalog.command(command).is_some();
                let desc = self
                    .catalog
                    .command(command)
                    .map(|c| c.description.chars().take(220).collect::<String>())
                    .unwrap_or_default();
                if ribbon_large(ui, ic, label, exists, false)
                    .on_hover_ui(|ui| {
                        ui.set_max_width(320.0);
                        ui.label(egui::RichText::new(*command).monospace().strong());
                        ui.label(desc.clone());
                    })
                    .clicked()
                {
                    self.cmds.push(Cmd::AddCommandWith(
                        command.to_string(),
                        name.to_string(),
                        values.iter().map(|(a, b)| (a.to_string(), b.to_string())).collect(),
                    ));
                }
            }
        });
    }

    fn ribbon_run(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "Deployment", |ui| {
            if ribbon_large(ui, icon::PLAY, "Install", true, true).on_hover_text("F5").clicked() {
                self.cmds.push(Cmd::Run(DeploymentType::Installation));
            }
            if ribbon_large(ui, icon::ARROW_COUNTER_CLOCKWISE, "Uninstall", true, false).clicked() {
                self.cmds.push(Cmd::Run(DeploymentType::Uninstall));
            }
            if ribbon_large(ui, icon::WRENCH, "Repair", true, false).clicked() {
                self.cmds.push(Cmd::Run(DeploymentType::Repair));
            }
        });
        ribbon_group(ui, "Isolated", |ui| {
            if ribbon_large(ui, icon::CUBE, "Windows Sandbox", true, false)
                .on_hover_text("Map the project read-only into Windows Sandbox and start the installation there")
                .clicked()
            {
                self.cmds.push(Cmd::Sandbox);
            }
        });
        ribbon_group(ui, "Remote", |ui| {
            if ribbon_large(ui, icon::DESKTOP_TOWER, "Run on Remote", true, false)
                .on_hover_text("Run on Remote Machine: copy the package to Windows computers and start the deployment there")
                .clicked()
            {
                self.cmds.push(Cmd::RemoteRun);
            }
        });
        ribbon_group(ui, "Options", |ui| {
            ribbon_stack(ui, |ui| {
                ribbon_toggle(ui, &mut self.prefs.silent_mode, "Silent mode");
                if ribbon_small(ui, icon::TERMINAL_WINDOW, "PowerShell Here", true).clicked() {
                    self.cmds.push(Cmd::PowerShellHere);
                }
                if ribbon_small(ui, icon::LIST_MAGNIFYING_GLASS, "Log Reader", true).clicked() {
                    self.cmds.push(Cmd::LogViewer);
                }
            });
        });
        ribbon_group(ui, "Logs", |ui| {
            if ribbon_large(ui, icon::FOLDER_OPEN, "Software Logs", true, false)
                .on_hover_text("C:\\Windows\\Logs\\Software")
                .clicked()
            {
                self.cmds.push(Cmd::OpenFolder(FolderKind::Logs));
            }
        });
    }

    fn ribbon_project(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "Project", |ui| {
            if ribbon_large(ui, icon::GEAR_SIX, "Settings", true, false).clicked() {
                self.cmds.push(Cmd::ProjectSettings);
            }
            ribbon_stack(ui, |ui| {
                if ribbon_small(ui, icon::FILE_PLUS, "New Project", true).clicked() {
                    self.cmds.push(Cmd::NewProject(NewKind::Blank));
                }
                if ribbon_small(ui, icon::FOLDER_OPEN, "Open Project", true).clicked() {
                    self.cmds.push(Cmd::BrowseOpenProject);
                }
                if ribbon_small(ui, icon::X_CIRCLE, "Close Project", true).clicked() {
                    self.cmds.push(Cmd::CloseProject);
                }
            });
        });
        ribbon_group(ui, "New From Installer", |ui| {
            if ribbon_large(ui, icon::DOWNLOAD_SIMPLE, "WinGet", true, false)
                .on_hover_text("Search the WinGet catalog and create a project from it")
                .clicked()
            {
                self.cmds.push(Cmd::WinGetImport);
            }
            if ribbon_large(ui, icon::DISC, "MSI", true, false).clicked() {
                self.cmds.push(Cmd::NewProject(NewKind::Msi(String::new())));
            }
            if ribbon_large(ui, icon::APP_WINDOW, "EXE", true, false).clicked() {
                self.cmds.push(Cmd::NewProject(NewKind::Exe(String::new())));
            }
        });
        ribbon_group(ui, "Folders", |ui| {
            if ribbon_large(ui, icon::FOLDER_OPEN, "Project Folder", true, false).clicked() {
                self.cmds.push(Cmd::OpenFolder(FolderKind::Project));
            }
            ribbon_stack(ui, |ui| {
                if ribbon_small(ui, icon::FOLDER, "Files", true).clicked() {
                    self.cmds.push(Cmd::OpenFolder(FolderKind::Files));
                }
                if ribbon_small(ui, icon::FOLDER, "SupportFiles", true).clicked() {
                    self.cmds.push(Cmd::OpenFolder(FolderKind::SupportFiles));
                }
                if ribbon_small(ui, icon::FOLDER, "Software Logs", true).clicked() {
                    self.cmds.push(Cmd::OpenFolder(FolderKind::Logs));
                }
            });
        });
        ribbon_group(ui, "Open Deployment File With", |ui| {
            if ribbon_large(ui, icon::CODE, "VS Code", true, false).clicked() {
                self.cmds.push(Cmd::OpenWith(Editor::VsCode));
            }
            ribbon_stack(ui, |ui| {
                if ribbon_small(ui, icon::TERMINAL_WINDOW, "PowerShell ISE", true).clicked() {
                    self.cmds.push(Cmd::OpenWith(Editor::Ise));
                }
                if ribbon_small(ui, icon::NOTE_PENCIL, "Notepad++", true).clicked() {
                    self.cmds.push(Cmd::OpenWith(Editor::NotepadPlusPlus));
                }
            });
        });
        ribbon_group(ui, "Export", |ui| {
            if ribbon_large(ui, icon::FILE_PDF, "Export PDF", true, false)
                .on_hover_text("Document the installation procedure as PDF")
                .clicked()
            {
                self.cmds.push(Cmd::ExportPdf);
            }
            if ribbon_large(ui, icon::PACKAGE, "Intune Package", true, false)
                .on_hover_text("Create an .intunewin file (built-in packager, no IntuneWinAppUtil needed)")
                .clicked()
            {
                self.cmds.push(Cmd::Intune);
            }
        });
    }

    fn ribbon_intune(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "Detection", |ui| {
            let has_rule = self.project.as_ref().map(|p| crate::detection::DetectionRule::load(p).is_some()).unwrap_or(false);
            let label = "Detection Rule";
            if ribbon_large(ui, icon::MAGNIFYING_GLASS, label, true, !has_rule)
                .on_hover_text("How Intune detects that the app is installed")
                .clicked()
            {
                self.cmds.push(Cmd::DetectionRule);
            }
            let errors = self.error_count();
            let label = if errors > 0 { format!("Check ({errors})") } else { "Check".to_string() };
            if ribbon_large(ui, icon::LIST_CHECKS, &label, true, false)
                .on_hover_text("Check the package for problems (F7)")
                .clicked()
            {
                self.cmds.push(Cmd::ShowProblems);
            }
        });
        ribbon_group(ui, "Publish", |ui| {
            if ribbon_large(ui, icon::CLOUD_ARROW_UP, "Upload to Intune", true, true)
                .on_hover_text("Package and upload the app to Microsoft Intune as Win32 app")
                .clicked()
            {
                self.cmds.push(Cmd::IntuneUpload);
            }
            if ribbon_large(ui, icon::PACKAGE, "Intune Package", true, false)
                .on_hover_text("Only create the .intunewin file")
                .clicked()
            {
                self.cmds.push(Cmd::Intune);
            }
        });
        ribbon_group(ui, "Documentation", |ui| {
            if ribbon_large(ui, icon::FILE_PDF, "Export PDF", true, false)
                .on_hover_text("Document the installation procedure as PDF")
                .clicked()
            {
                self.cmds.push(Cmd::ExportPdf);
            }
        });
    }

    fn ribbon_tools(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "Plugins", |ui| {
            if ribbon_large(ui, icon::PUZZLE_PIECE, "Plugins", true, false).clicked() {
                self.cmds.push(Cmd::ShowPlugins);
            }
            ribbon_stack(ui, |ui| {
                for (i, pl) in self.plugins.iter().take(3).enumerate() {
                    if ribbon_small(ui, icon::LIGHTNING, &pl.name, true).on_hover_text(&pl.description).clicked() {
                        self.cmds.push(Cmd::RunPlugin(i));
                    }
                }
            });
        });
        ribbon_group(ui, "Utilities", |ui| {
            if ribbon_large(ui, icon::LIST_MAGNIFYING_GLASS, "Log Reader", true, false)
                .on_hover_text("Read PSADT / CMTrace log files")
                .clicked()
            {
                self.cmds.push(Cmd::LogViewer);
            }
            if ribbon_large(ui, icon::LIST_CHECKS, "Executables List", true, false)
                .on_hover_text("Collect all .exe names below a folder (for -CloseProcesses / Block-AppExecution)")
                .clicked()
            {
                self.cmds.push(Cmd::ExecutablesList);
            }
            if ribbon_large(ui, icon::TERMINAL, "PSADT Help", true, false)
                .on_hover_text("Interactive PSADT command help console")
                .clicked()
            {
                self.cmds.push(Cmd::HelpConsole);
            }
        });
    }

    fn ribbon_view(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "Panes", |ui| {
            ribbon_stack(ui, |ui| {
                ribbon_toggle(ui, &mut self.prefs.show_library, "Command library");
                ribbon_toggle(ui, &mut self.prefs.show_inspector, "Properties");
                let mut dark = self.prefs.dark_mode;
                if ribbon_toggle(ui, &mut dark, "Dark mode").changed() {
                    self.cmds.push(Cmd::ToggleDark);
                }
            });
        });
        ribbon_group(ui, "Sequence", |ui| {
            if ribbon_large(ui, icon::ARROWS_OUT_SIMPLE, "Expand All", true, false).clicked() {
                self.cmds.push(Cmd::ExpandAll(true));
            }
            if ribbon_large(ui, icon::ARROWS_IN_SIMPLE, "Collapse All", true, false).clicked() {
                self.cmds.push(Cmd::ExpandAll(false));
            }
        });
        ribbon_group(ui, "Deployment Type", |ui| {
            for (dt, ic) in [
                (DeploymentType::Installation, icon::DOWNLOAD_SIMPLE),
                (DeploymentType::Uninstall, icon::TRASH_SIMPLE),
                (DeploymentType::Repair, icon::WRENCH),
            ] {
                let accent = self.dt == dt;
                if ribbon_large(ui, ic, dt.as_db(), true, accent).clicked() {
                    self.cmds.push(Cmd::SetDeploymentType(dt));
                }
            }
        });
    }

    fn ribbon_help(&mut self, ui: &mut egui::Ui) {
        ribbon_group(ui, "PSAppDeployToolkit", |ui| {
            if ribbon_large(ui, icon::BOOK_OPEN, "Documentation", true, false).clicked() {
                self.cmds.push(Cmd::OpenUrl("https://psappdeploytoolkit.com/docs/reference".into()));
            }
            if ribbon_large(ui, icon::CURRENCY_DOLLAR_SIMPLE, "Variables", true, false).clicked() {
                self.cmds.push(Cmd::OpenUrl("https://psappdeploytoolkit.com/docs/reference/variables".into()));
            }
            if ribbon_large(ui, icon::QUESTION, "Command Help", true, false).on_hover_text("F1").clicked() {
                self.cmds.push(Cmd::CommandHelp);
            }
        });
        ribbon_group(ui, "Deployment Editor", |ui| {
            if ribbon_large(ui, icon::GITHUB_LOGO, "GitHub", true, false).clicked() {
                self.cmds.push(Cmd::OpenUrl("https://github.com/tugich/DeploymentEditor".into()));
            }
            if ribbon_large(ui, icon::INFO, "About", true, false).clicked() {
                self.cmds.push(Cmd::Backstage(Some(crate::app::BackstagePage::About)));
            }
            if ribbon_large(ui, icon::HEART, "Donate", true, false).clicked() {
                self.cmds.push(Cmd::OpenUrl("https://www.paypal.com/donate/?hosted_button_id=PXABL8ESQQ4F8".into()));
            }
        });
    }
}
