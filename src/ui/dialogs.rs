//! Modal dialogs, the script preview window and the log reader.

use crate::widgets::segmented;
use eframe::egui::{self, CornerRadius, FontId, Stroke, Vec2};
use egui_phosphor::regular as icon;

use crate::app::{Cmd, DeploymentEditorApp, Dialog, NewKind, ToastKind};
use crate::logs::Severity;
use crate::theme::semibold;
use crate::widgets::{it, field_label, muted, primary_button, secondary_button};

enum After {
    Keep,
    Close,
    CreateProject,
    SaveSettings,
    SaveDetection,
    RemoveDetection,
}

impl DeploymentEditorApp {
    pub(super) fn dialogs(&mut self, ctx: &egui::Context) {
        let Some(mut dialog) = self.dialog.take() else { return };
        let p = self.palette;
        let mut after = After::Keep;

        let modal = egui::Modal::new(egui::Id::new("dialog")).frame(
            egui::Frame::new()
                .fill(p.panel_bg)
                .stroke(Stroke::new(1.0, p.border))
                .corner_radius(CornerRadius::same(12))
                .inner_margin(egui::Margin::same(22)),
        );
        let resp = modal.show(ctx, |ui| match &mut dialog {
            Dialog::Message { title, body, error } => {
                ui.set_max_width(520.0);
                ui.horizontal(|ui| {
                    let (ic, col) = if *error { (icon::WARNING_CIRCLE, p.danger) } else { (icon::INFO, p.info) };
                    ui.label(it(ic).size(26.0).color(col));
                    ui.label(egui::RichText::new(title.as_str()).font(semibold(17.0)));
                });
                ui.add_space(8.0);
                egui::ScrollArea::vertical().max_height(320.0).show(ui, |ui| {
                    ui.add(egui::Label::new(body.as_str()).wrap());
                });
                ui.add_space(14.0);
                ui.horizontal(|ui| {
                    if primary_button(ui, "  OK  ").clicked() {
                        after = After::Close;
                    }
                    if secondary_button(ui, &format!("{}  Copy", icon::COPY)).clicked() {
                        ui.ctx().copy_text(format!("{title}\n\n{body}"));
                    }
                });
            }
            Dialog::ConfirmRemove(id, name) => {
                ui.set_max_width(440.0);
                ui.label(egui::RichText::new("Remove action?").font(semibold(17.0)));
                ui.add_space(6.0);
                ui.label(format!("\"{name}\" and all its parameter values will be deleted from the project."));
                ui.add_space(16.0);
                ui.horizontal(|ui| {
                    let r = ui.add(
                        egui::Button::new(it(format!("{}  Remove", icon::TRASH)).color(egui::Color32::WHITE).font(semibold(13.5)))
                            .fill(p.danger)
                            .corner_radius(CornerRadius::same(6))
                            .min_size(Vec2::new(0.0, 30.0)),
                    );
                    if r.clicked() || ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                        self.cmds.push(Cmd::ConfirmedRemove(*id));
                        after = After::Close;
                    }
                    if secondary_button(ui, "Cancel").clicked() {
                        after = After::Close;
                    }
                });
            }
            Dialog::ExecutablesList(list) => {
                ui.set_width(560.0);
                ui.label(egui::RichText::new("Executables list").font(semibold(17.0)));
                muted(ui, "Comma-separated process names — e.g. for Show-ADTInstallationWelcome -CloseProcesses.");
                ui.add_space(8.0);
                ui.add(egui::TextEdit::multiline(list).font(egui::TextStyle::Monospace).desired_rows(6).desired_width(f32::INFINITY));
                ui.add_space(10.0);
                ui.horizontal(|ui| {
                    if primary_button(ui, &format!("{}  Copy", icon::COPY)).clicked() {
                        ui.ctx().copy_text(list.clone());
                        self.toast(ToastKind::Success, "Copied to clipboard");
                    }
                    if secondary_button(ui, "Close").clicked() {
                        after = After::Close;
                    }
                });
            }
            Dialog::Remote(form) => {
                use crate::remote::{Method, DEPLOY_MODES};
                ui.set_width(580.0);
                ui.horizontal(|ui| {
                    ui.label(it(icon::DESKTOP_TOWER).size(24.0).color(p.accent));
                    ui.label(egui::RichText::new("Run on Remote Machine").font(semibold(17.0)));
                });
                muted(ui, "Copies the package to Windows computers and starts the deployment there. Progress, exit codes and logs are shown in a PowerShell window.");
                ui.add_space(10.0);
                let o = &mut form.opts;
                ui.label(egui::RichText::new("Computers").font(semibold(13.0)));
                ui.add(
                    egui::TextEdit::multiline(&mut o.computers)
                        .hint_text("PC0815, PC0816.contoso.local — one or more names, separated by commas or new lines")
                        .desired_rows(2)
                        .desired_width(f32::INFINITY),
                );
                ui.add_space(8.0);
                egui::Grid::new("remote-grid").num_columns(2).spacing([14.0, 10.0]).show(ui, |ui| {
                    ui.label(egui::RichText::new("Deployment").font(semibold(13.0)));
                    use crate::project::DeploymentType;
                    let mut dt = o.deployment.unwrap_or(DeploymentType::Installation);
                    if segmented(
                        ui,
                        &mut dt,
                        &[
                            (DeploymentType::Installation, "Install".into()),
                            (DeploymentType::Uninstall, "Uninstall".into()),
                            (DeploymentType::Repair, "Repair".into()),
                        ],
                    ) {
                        o.deployment = Some(dt);
                    }
                    ui.end_row();
                    ui.label(egui::RichText::new("Deploy mode").font(semibold(13.0)));
                    let mut mode = DEPLOY_MODES.iter().position(|m| *m == o.deploy_mode).unwrap_or(0);
                    if segmented(ui, &mut mode, &[(0, "Silent".into()), (1, "NonInteractive".into()), (2, "Interactive".into())]) {
                        o.deploy_mode = DEPLOY_MODES[mode].to_string();
                    }
                    ui.end_row();
                    ui.label(egui::RichText::new("Method").font(semibold(13.0)));
                    segmented(
                        ui,
                        &mut o.method,
                        &[(Method::Auto, Method::Auto.label().into()), (Method::WinRm, Method::WinRm.label().into()), (Method::Task, Method::Task.label().into())],
                    );
                    ui.end_row();
                    ui.label("");
                    ui.add(egui::Label::new(egui::RichText::new(o.method.help()).size(11.5).color(p.text_muted)).wrap());
                    ui.end_row();
                    ui.label(egui::RichText::new("Account").font(semibold(13.0)));
                    ui.vertical(|ui| {
                        ui.radio_value(&mut o.ask_credentials, false, "My Windows account");
                        ui.radio_value(&mut o.ask_credentials, true, "Another admin account (PowerShell asks for the password)");
                    });
                    ui.end_row();
                });
                if o.method == Method::WinRm && o.deploy_mode == "Interactive" {
                    ui.label(
                        it(format!("{}  PowerShell Remoting has no desktop — dialogs can't be shown. Use the scheduled task for Interactive.", icon::WARNING))
                            .size(12.0)
                            .color(p.warning),
                    );
                }
                ui.add_space(6.0);
                egui::CollapsingHeader::new(egui::RichText::new("Advanced").font(semibold(13.0))).id_salt("remote-adv").show(ui, |ui| {
                    egui::Grid::new("remote-adv-grid").num_columns(2).spacing([14.0, 8.0]).show(ui, |ui| {
                        ui.label("Folder on the target");
                        ui.add(egui::TextEdit::singleline(&mut o.remote_dir).font(egui::TextStyle::Monospace).desired_width(330.0));
                        ui.end_row();
                        ui.label("Timeout");
                        ui.add(egui::DragValue::new(&mut o.timeout_minutes).range(1..=1440).suffix(" min"));
                        ui.end_row();
                        ui.label("");
                        ui.checkbox(&mut o.cleanup, "Delete the package on the target afterwards");
                        ui.end_row();
                        ui.label("");
                        ui.checkbox(&mut o.copy_logs, "Copy the PSADT logs back to this computer");
                        ui.end_row();
                    });
                });
                if let Some(e) = &form.error {
                    ui.add_space(6.0);
                    ui.label(it(format!("{}  {e}", icon::WARNING_CIRCLE)).size(12.5).color(p.danger));
                }
                ui.add_space(14.0);
                let n = crate::remote::computer_list(&form.opts.computers).map(|c| c.len()).unwrap_or(0);
                let label = match n {
                    0 | 1 => format!("{}  Run", icon::PLAY),
                    n => format!("{}  Run on {n} computers", icon::PLAY),
                };
                ui.horizontal(|ui| {
                    if primary_button(ui, &label).clicked() {
                        match crate::remote::computer_list(&form.opts.computers)
                            .and_then(|_| crate::remote::check_remote_dir(&form.opts.remote_dir))
                        {
                            Ok(_) => {
                                self.cmds.push(Cmd::RemoteRunNow(form.opts.clone()));
                                after = After::Close;
                            }
                            Err(e) => form.error = Some(format!("{e:#}")),
                        }
                    }
                    if secondary_button(ui, "Cancel").clicked() {
                        after = After::Close;
                    }
                });
            }
            Dialog::ExportPdf(opts) => {
                ui.set_width(460.0);
                ui.horizontal(|ui| {
                    ui.label(it(icon::FILE_PDF).size(24.0).color(p.accent));
                    ui.label(egui::RichText::new("Export documentation as PDF").font(semibold(17.0)));
                });
                ui.add_space(4.0);
                muted(ui, "Documents the package: overview, Intune settings, detection rule and every step of the install, uninstall and repair procedure.");
                ui.add_space(12.0);
                ui.checkbox(&mut opts.include_disabled, "Include disabled steps (marked as disabled)");
                ui.checkbox(&mut opts.include_checks, "Include the results of the quality check");
                ui.checkbox(&mut opts.include_code, "Appendix with the generated PowerShell code");
                ui.add_space(16.0);
                ui.horizontal(|ui| {
                    if primary_button(ui, &format!("{}  Export…", icon::EXPORT)).clicked() {
                        self.cmds.push(Cmd::ExportPdfNow(*opts));
                        after = After::Close;
                    }
                    if secondary_button(ui, "Cancel").clicked() {
                        after = After::Close;
                    }
                });
            }
            Dialog::ConfirmUnloadForPlugin(idx) => {
                ui.set_max_width(460.0);
                ui.label(egui::RichText::new("Unload project to continue").font(semibold(17.0)));
                ui.add_space(6.0);
                ui.label("To run this plugin, the project must be closed in the Deployment Editor. All changes are saved first.");
                ui.add_space(14.0);
                ui.horizontal(|ui| {
                    if primary_button(ui, "Save, close & run").clicked() {
                        let i = *idx;
                        self.commit_draft();
                        self.dialog = None;
                        self.run_plugin(i, true);
                        after = After::Close;
                    }
                    if secondary_button(ui, "Cancel").clicked() {
                        after = After::Close;
                    }
                });
            }
            Dialog::Plugins => {
                ui.set_width(640.0);
                ui.horizontal(|ui| {
                    ui.label(it(format!("{}  Plugins", icon::PUZZLE_PIECE)).font(semibold(18.0)));
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui.link("Open plugins folder").clicked() {
                            self.cmds.push(Cmd::OpenFolder(crate::app::FolderKind::Plugins));
                        }
                    });
                });
                muted(ui, "PowerShell scripts in Plugins\\Enabled-* receive -ProjectPath of the open project.");
                ui.add_space(10.0);
                egui::ScrollArea::vertical().max_height(380.0).show(ui, |ui| {
                    for (i, pl) in self.plugins.iter().enumerate() {
                        // The card itself is not clickable: an interaction over the whole
                        // card would sit on top of the Run button and swallow its clicks.
                        egui::Frame::new()
                            .fill(p.card_bg)
                            .stroke(Stroke::new(1.0, p.border))
                            .corner_radius(CornerRadius::same(10))
                            .inner_margin(egui::Margin::same(12))
                            .show(ui, |ui| {
                                ui.set_width(ui.available_width());
                                ui.horizontal(|ui| {
                                    ui.label(it(icon::LIGHTNING).size(20.0).color(p.accent));
                                    let text_width = (ui.available_width() - 96.0).max(120.0);
                                    ui.vertical(|ui| {
                                        ui.set_width(text_width);
                                        ui.horizontal(|ui| {
                                            ui.label(egui::RichText::new(&pl.name).font(semibold(14.0)));
                                            ui.label(it(format!("v{}", pl.version)).size(11.5).color(p.text_faint));
                                            if pl.unload_project {
                                                ui.label(egui::RichText::new("closes project").size(11.0).color(p.warning));
                                            }
                                        });
                                        ui.add(egui::Label::new(egui::RichText::new(&pl.description).size(12.5).color(p.text_muted)).wrap());
                                        ui.label(it(format!("by {}  ·  {}", pl.author, pl.date)).size(11.0).color(p.text_faint));
                                    });
                                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                        if primary_button(ui, &format!("{}  Run", icon::PLAY)).clicked() {
                                            self.cmds.push(Cmd::RunPlugin(i));
                                            after = After::Close;
                                        }
                                    });
                                });
                            });
                        ui.add_space(6.0);
                    }
                    if self.plugins.is_empty() {
                        muted(ui, "No enabled plugins found.");
                    }
                });
                ui.add_space(10.0);
                if secondary_button(ui, "Close").clicked() {
                    after = After::Close;
                }
            }
            Dialog::Problems { then } => {
                ui.set_width(680.0);
                let issues = self.all_issues();
                let errors = crate::validate::count(&issues, crate::validate::Severity::Error);
                let warnings = crate::validate::count(&issues, crate::validate::Severity::Warning);
                let (ic, col, title) = if errors > 0 {
                    (icon::WARNING_CIRCLE, p.danger, format!("{errors} error{} found", if errors == 1 { "" } else { "s" }))
                } else if warnings > 0 {
                    (icon::WARNING, p.warning, format!("No errors · {warnings} warning{}", if warnings == 1 { "" } else { "s" }))
                } else {
                    (icon::CHECK_CIRCLE, p.success, "No problems found".to_string())
                };
                ui.horizontal(|ui| {
                    ui.label(it(ic).size(24.0).color(col));
                    ui.label(egui::RichText::new(title).font(semibold(18.0)));
                    if self.ps_check_running || self.ps_check_due.is_some() {
                        ui.add(egui::Spinner::new().size(14.0));
                        muted(ui, "PowerShell syntax check…");
                    }
                });
                if then.is_some() && errors > 0 {
                    muted(ui, "These errors would very likely break the deployment. Fix them first, or continue anyway.");
                } else {
                    muted(ui, "Click a problem to jump to the step.");
                }
                if let Some(e) = &self.ps_check_error {
                    ui.add_space(4.0);
                    ui.label(
                        it(format!("{}  The PowerShell syntax check could not run: {e}", icon::INFO))
                            .size(12.0)
                            .color(p.warning),
                    );
                }
                ui.add_space(8.0);
                egui::ScrollArea::vertical().max_height(420.0).show(ui, |ui| {
                    for issue in &issues {
                        let (ic, col) = match issue.severity {
                            crate::validate::Severity::Error => (icon::X_CIRCLE, p.danger),
                            crate::validate::Severity::Warning => (icon::WARNING, p.warning),
                            crate::validate::Severity::Hint => (icon::INFO, p.info),
                        };
                        let resp = egui::Frame::new()
                            .fill(p.card_bg)
                            .stroke(Stroke::new(1.0, p.border))
                            .corner_radius(CornerRadius::same(8))
                            .inner_margin(egui::Margin::symmetric(10, 7))
                            .show(ui, |ui| {
                                ui.set_width(ui.available_width());
                                ui.horizontal(|ui| {
                                    ui.label(it(ic).size(17.0).color(col));
                                    ui.vertical(|ui| {
                                        ui.label(egui::RichText::new(issue.location()).size(11.5).color(p.text_muted));
                                        ui.add(egui::Label::new(egui::RichText::new(&issue.message).size(13.0)).wrap());
                                    });
                                });
                            })
                            .response;
                        // The row has no buttons inside, so a whole-row click area is safe here.
                        let resp = ui.interact(resp.rect, ui.id().with(("issue", &issue.message, issue.action_id)), egui::Sense::click());
                        if resp.hovered() {
                            ui.painter().rect_stroke(resp.rect, CornerRadius::same(8), Stroke::new(1.5, p.accent), egui::StrokeKind::Inside);
                        }
                        if resp.on_hover_cursor(egui::CursorIcon::PointingHand).clicked() {
                            if let Some(dt) = issue.dt {
                                self.cmds.push(Cmd::GoTo(dt, issue.action_id));
                            } else if issue.message.contains("Project settings") {
                                self.cmds.push(Cmd::ProjectSettings);
                            } else if issue.message.contains("etection rule") {
                                self.cmds.push(Cmd::DetectionRule);
                            }
                            after = After::Close;
                        }
                        ui.add_space(4.0);
                    }
                });
                ui.add_space(12.0);
                ui.horizontal(|ui| {
                    match then {
                        Some(cmd) if errors > 0 => {
                            if primary_button(ui, &format!("{}  Fix problems", icon::WRENCH)).clicked() {
                                if let Some(first) = issues.iter().find(|i| i.severity == crate::validate::Severity::Error) {
                                    if let Some(dt) = first.dt {
                                        self.cmds.push(Cmd::GoTo(dt, first.action_id));
                                    }
                                }
                                after = After::Close;
                            }
                            if secondary_button(ui, "Continue anyway").clicked() {
                                self.gate_override = true;
                                self.cmds.push((**cmd).clone());
                                after = After::Close;
                            }
                        }
                        Some(cmd) => {
                            // Errors were fixed while the dialog was open.
                            if primary_button(ui, &format!("{}  Continue", icon::CHECK)).clicked() {
                                self.cmds.push((**cmd).clone());
                                after = After::Close;
                            }
                        }
                        None => {
                            if primary_button(ui, "Close").clicked() {
                                after = After::Close;
                            }
                        }
                    }
                    if secondary_button(ui, &format!("{}  Check again", icon::ARROWS_CLOCKWISE)).clicked() {
                        self.revalidate();
                        if cfg!(windows) {
                            self.ps_check_due = Some(std::time::Instant::now());
                        }
                    }
                });
            }
            Dialog::Detection(f) => match self.detection_form(ui, f) {
                super::detection_dialog::Outcome::None => {}
                super::detection_dialog::Outcome::Save => after = After::SaveDetection,
                super::detection_dialog::Outcome::Remove => after = After::RemoveDetection,
                super::detection_dialog::Outcome::Close => after = After::Close,
            },
            Dialog::WinGet(f) => {
                ui.set_width(680.0);
                ui.label(it(format!("{}  Import from WinGet", icon::DOWNLOAD_SIMPLE)).font(semibold(18.0)));
                muted(ui, "Search the Windows Package Manager catalog — the installer is downloaded into Files and the install step is created for you.");
                ui.add_space(10.0);
                let mut do_search = false;
                ui.horizontal(|ui| {
                    let r = ui.add(
                        egui::TextEdit::singleline(&mut f.query)
                            .hint_text(it(format!("{}  e.g. 7zip, Notepad++, VLC", icon::MAGNIFYING_GLASS)))
                            .desired_width(520.0)
                            .margin(egui::Margin::symmetric(10, 6)),
                    );
                    if r.lost_focus() && ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                        do_search = true;
                    }
                    ui.add_enabled_ui(!f.searching && !f.query.trim().is_empty(), |ui| {
                        if primary_button(ui, "Search").clicked() {
                            do_search = true;
                        }
                    });
                    if f.searching {
                        ui.add(egui::Spinner::new());
                    }
                });
                if do_search && !f.query.trim().is_empty() {
                    f.searching = true;
                    f.error = None;
                    self.winget_search_async(f.query.trim().to_string());
                }
                ui.add_space(6.0);
                let mut load: Option<(String, String)> = None;
                if !f.results.is_empty() {
                    egui::Frame::new()
                        .stroke(Stroke::new(1.0, p.border))
                        .corner_radius(CornerRadius::same(8))
                        .inner_margin(egui::Margin::same(4))
                        .show(ui, |ui| {
                            egui::ScrollArea::vertical().id_salt("wg-results").max_height(170.0).show(ui, |ui| {
                                ui.set_width(ui.available_width());
                                for (i, h) in f.results.iter().enumerate() {
                                    let sel = f.selected == Some(i);
                                    let (rect, r) = ui.allocate_exact_size(Vec2::new(ui.available_width(), 26.0), egui::Sense::click());
                                    if sel {
                                        ui.painter().rect_filled(rect, CornerRadius::same(5), p.accent_soft);
                                    } else if r.hovered() {
                                        ui.painter().rect_filled(rect, CornerRadius::same(5), p.hover);
                                    }
                                    let y = rect.center().y;
                                    ui.painter().text(egui::pos2(rect.left() + 10.0, y), egui::Align2::LEFT_CENTER, &h.name, semibold(13.0), p.text);
                                    ui.painter().text(egui::pos2(rect.left() + 230.0, y), egui::Align2::LEFT_CENTER, &h.id, FontId::monospace(12.0), p.text_muted);
                                    ui.painter().text(egui::pos2(rect.right() - 10.0, y), egui::Align2::RIGHT_CENTER, &h.version, FontId::proportional(12.5), p.text_muted);
                                    if r.clicked() && !sel {
                                        f.selected = Some(i);
                                        f.id = h.id.clone();
                                        f.version = h.version.clone();
                                        load = Some((h.id.clone(), h.version.clone()));
                                    }
                                }
                            });
                        });
                    ui.add_space(6.0);
                }
                ui.horizontal(|ui| {
                    ui.label(it("Package").color(p.text_muted));
                    ui.add(egui::TextEdit::singleline(&mut f.id).hint_text("Identifier").font(egui::TextStyle::Monospace).desired_width(260.0));
                    ui.add(egui::TextEdit::singleline(&mut f.version).hint_text("Version").font(egui::TextStyle::Monospace).desired_width(110.0));
                    ui.add_enabled_ui(!f.id.trim().is_empty() && !f.version.trim().is_empty() && !f.loading_manifest, |ui| {
                        if secondary_button(ui, "Load manifest").clicked() {
                            load = Some((f.id.clone(), f.version.clone()));
                        }
                    });
                });
                if let Some((id, ver)) = load {
                    f.loading_manifest = true;
                    f.manifest = None;
                    f.installer = None;
                    f.error = None;
                    if f.folder.trim().is_empty() || f.folder.contains("(WinGet)") {
                        let base = dirs_documents().join("PSADT Projects");
                        f.folder = base.join(format!("{} {} (WinGet)", id, ver)).display().to_string();
                    }
                    self.winget_manifest_async(id, ver);
                }
                if f.loading_manifest {
                    ui.horizontal(|ui| {
                        ui.add(egui::Spinner::new());
                        muted(ui, "Loading the installer manifest from winget-pkgs…");
                    });
                }
                if let Some(m) = &f.manifest {
                    ui.add_space(6.0);
                    ui.horizontal(|ui| {
                        field_label(ui, "Installer", None);
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            let before = f.arch.clone();
                            egui::ComboBox::from_id_salt("wg-arch").selected_text(f.arch.clone()).width(90.0).show_ui(ui, |ui| {
                                for a in ["x64", "x86", "arm64", "neutral"] {
                                    ui.selectable_value(&mut f.arch, a.to_string(), a);
                                }
                            });
                            ui.label(it("Target architecture").size(12.0).color(p.text_muted));
                            if f.arch != before {
                                f.installer = m.best_for(&f.arch);
                            }
                        });
                    });
                    egui::ScrollArea::vertical().id_salt("wg-inst").max_height(150.0).show(ui, |ui| {
                        for (i, ins) in m.installers.iter().enumerate() {
                            let supported = ins.local_extension().is_some();
                            ui.add_enabled_ui(supported, |ui| {
                                ui.horizontal(|ui| {
                                    ui.radio_value(&mut f.installer, Some(i), it(ins.label()).font(semibold(12.5)));
                                    let sw = ins.silent_switch();
                                    let info = if ins.is_msi() { "msi".to_string() } else if sw.is_empty() { "no silent switch".into() } else { sw };
                                    crate::widgets::chip(ui, &info, p.text_muted, p.hover);
                                });
                                ui.label(it(&ins.url).size(11.5).color(p.text_faint));
                            });
                        }
                    });
                    ui.add_space(6.0);
                    field_label(ui, "Project folder", Some("the PSADT framework is copied here"));
                    ui.horizontal(|ui| {
                        ui.add(egui::TextEdit::singleline(&mut f.folder).desired_width(550.0).margin(egui::Margin::symmetric(8, 6)));
                        if secondary_button(ui, "Browse…").clicked() {
                            if let Some(d) = rfd::FileDialog::new().set_title("Folder for the new project").pick_folder() {
                                f.folder = d.display().to_string();
                            }
                        }
                    });
                    let not_empty = std::path::Path::new(f.folder.trim()).read_dir().map(|mut d| d.next().is_some()).unwrap_or(false);
                    if not_empty {
                        ui.checkbox(&mut f.confirm_overwrite, it(format!("{}  Folder is not empty — overwrite template files", icon::WARNING)).color(p.warning));
                    }
                    ui.add_space(10.0);
                    let ok = f.installer.is_some() && !f.folder.trim().is_empty() && (!not_empty || f.confirm_overwrite) && !f.creating;
                    ui.horizontal(|ui| {
                        ui.add_enabled_ui(ok, |ui| {
                            if primary_button(ui, &format!("{}  Download & create project", icon::DOWNLOAD_SIMPLE)).clicked() {
                                f.creating = true;
                                f.error = None;
                                self.winget_create_async(std::path::PathBuf::from(f.folder.trim()), m.clone(), f.installer.unwrap_or(0));
                            }
                        });
                        if f.creating {
                            ui.add(egui::Spinner::new());
                            muted(ui, "Downloading installer…");
                        }
                    });
                }
                if let Some(e) = &f.error {
                    ui.add_space(6.0);
                    ui.add(egui::Label::new(it(e.as_str()).color(p.danger)).wrap());
                }
                ui.add_space(8.0);
                if secondary_button(ui, "Close").clicked() {
                    after = After::Close;
                }
            }
            Dialog::ProjectSettings(s) => {
                ui.set_width(520.0);
                ui.label(it(format!("{}  Project settings", icon::GEAR_SIX)).font(semibold(18.0)));
                muted(ui, "Written into the $adtSession block of the deployment script.");
                ui.add_space(12.0);
                let w = 476.0;
                field_label(ui, "Project name", None);
                ui.add(egui::TextEdit::singleline(&mut s.project_name).desired_width(w).margin(egui::Margin::symmetric(8, 6)));
                ui.add_space(4.0);
                ui.columns(2, |cols| {
                    field_label(&mut cols[0], "Application name", Some("AppName"));
                    cols[0].add(egui::TextEdit::singleline(&mut s.app_name).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)));
                    field_label(&mut cols[1], "Vendor", Some("AppVendor"));
                    cols[1].add(egui::TextEdit::singleline(&mut s.app_vendor).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)));
                    field_label(&mut cols[0], "Version", Some("AppVersion"));
                    cols[0].add(egui::TextEdit::singleline(&mut s.app_version).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)));
                    field_label(&mut cols[1], "Language", Some("AppLang"));
                    cols[1].add(egui::TextEdit::singleline(&mut s.app_language).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)));
                    field_label(&mut cols[0], "Architecture", Some("AppArch"));
                    egui::ComboBox::from_id_salt("arch")
                        .selected_text(if s.app_architecture.is_empty() { "—".to_string() } else { s.app_architecture.clone() })
                        .width(cols[0].available_width())
                        .show_ui(&mut cols[0], |ui| {
                            for a in ["x64", "x86", "arm64", ""] {
                                ui.selectable_value(&mut s.app_architecture, a.to_string(), if a.is_empty() { "(any)" } else { a });
                            }
                        });
                    field_label(&mut cols[1], "Script author", Some("AppScriptAuthor"));
                    cols[1].add(egui::TextEdit::singleline(&mut s.app_author).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)));
                });
                ui.add_space(16.0);
                ui.horizontal(|ui| {
                    if primary_button(ui, &format!("{}  Save", icon::FLOPPY_DISK)).clicked() {
                        after = After::SaveSettings;
                    }
                    if secondary_button(ui, "Cancel").clicked() {
                        after = After::Close;
                    }
                });
            }
            Dialog::NewProject(form) => {
                ui.set_width(560.0);
                let (ic, title) = match &form.kind {
                    NewKind::Blank => (icon::FILE_PLUS, "New project"),
                    NewKind::Exe(_) => (icon::APP_WINDOW, "Import executable"),
                    NewKind::Msi(_) => (icon::DISC, "Import Windows Installer (MSI)"),
                };
                ui.label(it(format!("{ic}  {title}")).font(semibold(18.0)));
                ui.add_space(12.0);

                // Installer picker
                let mut pick_msi = None;
                let is_msi = matches!(form.kind, NewKind::Msi(_));
                if let NewKind::Exe(path_ref) | NewKind::Msi(path_ref) = &mut form.kind {
                    field_label(ui, "Installer", None);
                    let folder = &mut form.folder;
                    ui.horizontal(|ui| {
                        let r = ui.add(
                            egui::TextEdit::singleline(path_ref)
                                .hint_text(if is_msi { "C:\\Setup\\product.msi" } else { "C:\\Setup\\setup.exe" })
                                .desired_width(430.0)
                                .margin(egui::Margin::symmetric(8, 6)),
                        );
                        if is_msi && r.lost_focus() && std::path::Path::new(path_ref.trim()).is_file() {
                            pick_msi = Some(path_ref.trim().to_string());
                        }
                        if secondary_button(ui, "Browse…").clicked() {
                            let ext: &[&str] = if is_msi { &["msi"] } else { &["exe"] };
                            if let Some(f) = rfd::FileDialog::new()
                                .add_filter(if is_msi { "Windows Installer" } else { "Executable" }, ext)
                                .pick_file()
                            {
                                *path_ref = f.display().to_string();
                                if is_msi {
                                    pick_msi = Some(path_ref.clone());
                                }
                                if folder.is_empty() {
                                    if let (Some(stem), Some(parent)) = (f.file_stem(), f.parent()) {
                                        *folder = parent.join(format!("{} (PSADT)", stem.to_string_lossy())).display().to_string();
                                    }
                                }
                            }
                        }
                    });
                    if is_msi {
                        ui.add_space(4.0);
                        if form.msi_loading {
                            ui.horizontal(|ui| {
                                ui.add(egui::Spinner::new());
                                muted(ui, "Reading MSI properties…");
                            });
                        } else if let Some(info) = &form.msi {
                            egui::Frame::new().fill(p.accent_soft).corner_radius(CornerRadius::same(8)).inner_margin(egui::Margin::same(10)).show(ui, |ui| {
                                ui.set_width(ui.available_width());
                                egui::Grid::new("msi").num_columns(2).spacing(Vec2::new(14.0, 3.0)).show(ui, |ui| {
                                    for (k, v) in [("Product", &info.product_name), ("Version", &info.product_version), ("Manufacturer", &info.manufacturer), ("ProductCode", &info.product_code)] {
                                        ui.label(egui::RichText::new(k).color(p.text_muted));
                                        ui.label(egui::RichText::new(v).monospace());
                                        ui.end_row();
                                    }
                                });
                            });
                        } else {
                            muted(ui, "Product name, version and code are read with Scripts\\Read-Msi.ps1 after you pick a file.");
                        }
                    }
                    ui.add_space(6.0);
                }
                if let Some(path) = pick_msi {
                    form.msi_loading = true;
                    form.msi = None;
                    self.read_msi_async(path);
                }

                field_label(ui, "Project name", None);
                ui.add(egui::TextEdit::singleline(&mut form.name).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)));
                ui.add_space(6.0);
                field_label(ui, "Project folder", Some("the PSADT framework is copied here"));
                ui.horizontal(|ui| {
                    ui.add(egui::TextEdit::singleline(&mut form.folder).desired_width(430.0).margin(egui::Margin::symmetric(8, 6)));
                    if secondary_button(ui, "Browse…").clicked() {
                        if let Some(f) = rfd::FileDialog::new().set_title("Folder for the new project").pick_folder() {
                            form.folder = f.display().to_string();
                        }
                    }
                });
                let folder = std::path::Path::new(form.folder.trim());
                let not_empty = !form.folder.trim().is_empty()
                    && folder.read_dir().map(|mut d| d.next().is_some()).unwrap_or(false);
                if not_empty {
                    ui.add_space(4.0);
                    egui::Frame::new().fill(p.warning.gamma_multiply(0.12)).corner_radius(CornerRadius::same(8)).inner_margin(egui::Margin::same(10)).show(ui, |ui| {
                        ui.set_width(ui.available_width());
                        ui.label(it(format!("{}  The folder is not empty — existing template files will be overwritten.", icon::WARNING)).color(p.warning));
                        ui.checkbox(&mut form.confirm_overwrite, "I understand, overwrite files");
                    });
                }
                if let Some(e) = &form.error {
                    ui.add_space(4.0);
                    ui.label(egui::RichText::new(e.as_str()).color(p.danger));
                }
                ui.add_space(16.0);
                let installer_ok = match &form.kind {
                    NewKind::Blank => true,
                    NewKind::Exe(p) | NewKind::Msi(p) => std::path::Path::new(p).is_file(),
                };
                let ok = installer_ok && !form.folder.trim().is_empty() && (!not_empty || form.confirm_overwrite) && !form.msi_loading;
                ui.horizontal(|ui| {
                    ui.add_enabled_ui(ok, |ui| {
                        if primary_button(ui, &format!("{}  Create project", icon::CHECK)).clicked() {
                            after = After::CreateProject;
                        }
                    });
                    if secondary_button(ui, "Cancel").clicked() {
                        after = After::Close;
                    }
                });
            }
        });

        if resp.should_close() && matches!(after, After::Keep) {
            // Escape / click outside closes simple dialogs, but not forms with input.
            let escape = !resp.backdrop_response.clicked();
            if !matches!(dialog, Dialog::NewProject(_) | Dialog::ProjectSettings(_) | Dialog::WinGet(_) | Dialog::Detection(_)) || escape {
                after = After::Close;
            }
        }

        match after {
            After::Keep => {
                if self.dialog.is_none() {
                    self.dialog = Some(dialog);
                }
            }
            After::Close => {}
            After::SaveSettings => {
                if let Dialog::ProjectSettings(s) = &dialog {
                    self.checkpoint("Project settings");
                    if let Some(proj) = self.project.as_mut() {
                        match proj.save_settings(s) {
                            Ok(()) => {
                                self.project_name = proj.display_name();
                                self.toast(ToastKind::Success, "Project settings saved");
                            }
                            Err(e) => self.error("Can't save settings", format!("{e:#}")),
                        }
                    }
                }
            }
            After::SaveDetection => {
                if let Dialog::Detection(f) = &dialog {
                    self.save_detection(f.rule.as_ref());
                }
            }
            After::RemoveDetection => self.save_detection(None),
            After::CreateProject => {
                if let Dialog::NewProject(mut form) = dialog {
                    if let Err(e) = self.create_project(&form) {
                        form.error = Some(format!("{e:#}"));
                        self.dialog = Some(Dialog::NewProject(form));
                    }
                }
            }
        }
    }

    pub(super) fn preview_window(&mut self, ctx: &egui::Context) {
        let Some(preview) = self.preview.as_mut() else { return };
        let p = self.palette;
        let mut open = true;
        let mut switch_full: Option<bool> = None;
        let mut generate = false;
        let title = format!("{}  Preview — Invoke-AppDeployToolkit.ps1", icon::EYE);
        egui::Window::new(it(title))
            .open(&mut open)
            .pivot(egui::Align2::CENTER_CENTER)
            .default_pos(ctx.content_rect().center())
            .default_size(Vec2::new(900.0, 640.0))
            .collapsible(false)
            .resizable(true)
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    let mut full = preview.full;
                    let items = [(true, "Full script".to_string()), (false, format!("{} part only", self.dt.as_db()))];
                    if crate::widgets::segmented(ui, &mut full, &items) {
                        switch_full = Some(full);
                    }
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if primary_button(ui, &format!("{}  Generate file", icon::FILE_ARROW_DOWN)).clicked() {
                            generate = true;
                        }
                        if secondary_button(ui, &format!("{}  Copy", icon::COPY)).clicked() {
                            ui.ctx().copy_text(preview.text.clone());
                        }
                    });
                });
                ui.add_space(6.0);
                egui::Frame::new()
                    .fill(ui.visuals().code_bg_color)
                    .stroke(Stroke::new(1.0, p.border))
                    .corner_radius(CornerRadius::same(8))
                    .inner_margin(egui::Margin::same(10))
                    .show(ui, |ui| {
                        egui::ScrollArea::both().auto_shrink([false, false]).show(ui, |ui| {
                            let job = crate::highlight::highlight_powershell(&preview.text, FontId::monospace(12.5), p.dark);
                            let galley = ui.fonts_mut(|f| f.layout_job(job));
                            ui.add(egui::Label::new(galley).selectable(true));
                        });
                    });
            });
        if let Some(full) = switch_full {
            self.build_preview(full);
        }
        if generate {
            self.cmds.push(Cmd::Generate);
        }
        if !open {
            self.preview = None;
        }
    }

    pub(super) fn log_viewer_window(&mut self, ctx: &egui::Context) {
        let Some(lv) = self.log_viewer.as_mut() else { return };
        let p = self.palette;
        let mut open = true;
        let mut add_files: Option<Vec<std::path::PathBuf>> = None;
        egui::Window::new(it(format!("{}  Software Log Reader", icon::LIST_MAGNIFYING_GLASS)))
            .open(&mut open)
            .pivot(egui::Align2::CENTER_CENTER)
            .default_pos(ctx.content_rect().center())
            .default_size(Vec2::new(1000.0, 600.0))
            .collapsible(false)
            .resizable(true)
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    if primary_button(ui, &format!("{}  Open log files…", icon::FOLDER_OPEN)).clicked() {
                        add_files = rfd::FileDialog::new()
                            .set_directory(crate::logs::default_log_folder())
                            .add_filter("Log files", &["log", "txt", "lo_"])
                            .add_filter("All files", &["*"])
                            .pick_files();
                    }
                    if secondary_button(ui, &format!("{}  Reload", icon::ARROWS_CLOCKWISE)).clicked() {
                        lv.logs.reload();
                    }
                    if secondary_button(ui, "Clear").clicked() {
                        lv.logs = Default::default();
                        lv.selected = None;
                    }
                    ui.add_space(10.0);
                    ui.add(egui::TextEdit::singleline(&mut lv.filter).hint_text(it(format!("{}  Filter messages", icon::FUNNEL))).desired_width(240.0));
                    ui.toggle_value(&mut lv.show_info, "Info");
                    ui.toggle_value(&mut lv.show_warning, egui::RichText::new("Warning").color(p.warning));
                    ui.toggle_value(&mut lv.show_error, egui::RichText::new("Error").color(p.danger));
                });
                if !lv.logs.files.is_empty() {
                    ui.horizontal_wrapped(|ui| {
                        for f in &lv.logs.files {
                            crate::widgets::chip(ui, &f.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default(), p.text_muted, p.hover);
                        }
                    });
                }
                ui.add_space(4.0);
                let q = lv.filter.to_lowercase();
                let rows: Vec<usize> = lv
                    .logs
                    .entries
                    .iter()
                    .enumerate()
                    .filter(|(_, e)| match e.severity {
                        Severity::Info => lv.show_info,
                        Severity::Warning => lv.show_warning,
                        Severity::Error => lv.show_error,
                    })
                    .filter(|(_, e)| q.is_empty() || e.message.to_lowercase().contains(&q) || e.component.to_lowercase().contains(&q))
                    .map(|(i, _)| i)
                    .collect();
                if lv.logs.files.is_empty() {
                    ui.add_space(40.0);
                    ui.vertical_centered(|ui| {
                        ui.label(it(icon::FILE_TEXT).size(40.0).color(p.text_faint));
                        muted(ui, "Open one or more PSADT / CMTrace logs (C:\\Windows\\Logs\\Software).");
                    });
                    return;
                }
                let detail_h = if lv.selected.is_some() { 110.0 } else { 0.0 };
                let avail = ui.available_height() - detail_h;
                egui::ScrollArea::vertical().max_height(avail).auto_shrink([false, false]).show_rows(ui, 22.0, rows.len(), |ui, range| {
                    for &i in &rows[range] {
                        let e = &lv.logs.entries[i];
                        let (col, ic) = match e.severity {
                            Severity::Info => (p.text, icon::INFO),
                            Severity::Warning => (p.warning, icon::WARNING),
                            Severity::Error => (p.danger, icon::X_CIRCLE),
                        };
                        let sel = lv.selected == Some(i);
                        let (rect, resp) = ui.allocate_exact_size(Vec2::new(ui.available_width(), 22.0), egui::Sense::click());
                        let bg = if sel {
                            p.accent_soft
                        } else if e.severity == Severity::Error {
                            p.danger.gamma_multiply(0.09)
                        } else if e.severity == Severity::Warning {
                            p.warning.gamma_multiply(0.09)
                        } else if resp.hovered() {
                            p.hover
                        } else {
                            egui::Color32::TRANSPARENT
                        };
                        ui.painter().rect_filled(rect, CornerRadius::same(3), bg);
                        let y = rect.center().y;
                        let f = FontId::proportional(12.5);
                        ui.painter().text(egui::pos2(rect.left() + 10.0, y), egui::Align2::CENTER_CENTER, ic, crate::theme::icon_font(13.0), col);
                        ui.painter().text(egui::pos2(rect.left() + 24.0, y), egui::Align2::LEFT_CENTER, format!("{} {}", e.date, e.time), FontId::monospace(11.5), p.text_muted);
                        ui.painter().text(egui::pos2(rect.left() + 210.0, y), egui::Align2::LEFT_CENTER, &e.component, f.clone(), p.text_muted);
                        let first = e.message.lines().next().unwrap_or("");
                        let galley = ui.painter().layout(first.to_string(), f, col, (rect.width() - 360.0).max(50.0));
                        let clip = ui.painter().with_clip_rect(rect);
                        clip.galley(egui::pos2(rect.left() + 350.0, y - 8.0), galley, col);
                        if resp.clicked() {
                            lv.selected = Some(i);
                        }
                    }
                });
                if let Some(i) = lv.selected {
                    if let Some(e) = lv.logs.entries.get(i) {
                        ui.separator();
                        egui::ScrollArea::vertical().id_salt("logdetail").max_height(100.0).show(ui, |ui| {
                            let file = lv.logs.files.get(e.file).and_then(|f| f.file_name()).map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
                            ui.label(it(format!("{} {}  ·  {}  ·  {}", e.date, e.time, e.component, file)).size(11.5).color(p.text_muted));
                            ui.add(egui::Label::new(egui::RichText::new(&e.message).monospace()).wrap().selectable(true));
                        });
                    }
                }
            });
        if let Some(files) = add_files {
            for f in files {
                if let Err(e) = lv.logs.add_file(f.clone()) {
                    self.toasts.push(crate::app::Toast {
                        text: format!("Can't read {}: {e}", f.display()),
                        kind: ToastKind::Error,
                        created: std::time::Instant::now(),
                    });
                }
            }
        }
        if !open {
            self.log_viewer = None;
        }
    }
}

/// The user's Documents folder (falls back to the home directory).
fn dirs_documents() -> std::path::PathBuf {
    let home = std::env::var_os("USERPROFILE").or_else(|| std::env::var_os("HOME")).map(std::path::PathBuf::from).unwrap_or_default();
    let docs = home.join("Documents");
    if docs.is_dir() { docs } else { home }
}
