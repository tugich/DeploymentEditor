//! The "File" backstage view (start screen): new, open, recent, info, about.

use eframe::egui::{self, Align2, CornerRadius, FontId, Sense, Stroke, StrokeKind, Vec2};
use egui_phosphor::regular as icon;

use crate::app::{BackstagePage, Cmd, DeploymentEditorApp, NewKind};
use crate::project::DATABASE_FILE;
use crate::theme::semibold;
use crate::widgets::{it, muted, primary_button, secondary_button};

impl DeploymentEditorApp {
    pub(super) fn backstage(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        let page = self.backstage.unwrap_or(BackstagePage::Home);

        egui::Panel::left("backstage-nav")
            .exact_size(232.0)
            .resizable(false)
            .show_separator_line(false)
            .frame(egui::Frame::new().fill(p.accent).inner_margin(egui::Margin::symmetric(10, 16)))
            .show(ui, |ui| {
                ui.spacing_mut().item_spacing.y = 2.0;
                if self.project.is_some() {
                    if nav_item(ui, icon::ARROW_LEFT, "Back to project", false).clicked() {
                        self.cmds.push(Cmd::Backstage(None));
                    }
                    ui.add_space(10.0);
                } else {
                    ui.horizontal(|ui| {
                        ui.add_space(6.0);
                        ui.label(it(icon::PACKAGE).size(26.0).color(p.on_accent));
                        ui.vertical(|ui| {
                            ui.spacing_mut().item_spacing.y = 0.0;
                            ui.label(egui::RichText::new("Deployment").font(semibold(15.0)).color(p.on_accent));
                            ui.label(egui::RichText::new("Editor").font(semibold(15.0)).color(p.on_accent));
                        });
                    });
                    ui.add_space(18.0);
                }
                for (pg, ic, label) in [
                    (BackstagePage::Home, icon::HOUSE, "Home"),
                    (BackstagePage::New, icon::FILE_PLUS, "New"),
                    (BackstagePage::Open, icon::FOLDER_OPEN, "Open"),
                ] {
                    if nav_item(ui, ic, label, page == pg).clicked() {
                        self.cmds.push(Cmd::Backstage(Some(pg)));
                    }
                }
                if self.project.is_some() {
                    if nav_item(ui, icon::INFO, "Info", page == BackstagePage::Info).clicked() {
                        self.cmds.push(Cmd::Backstage(Some(BackstagePage::Info)));
                    }
                    if nav_item(ui, icon::FLOPPY_DISK, "Save", false).clicked() {
                        self.cmds.push(Cmd::Save);
                    }
                    if nav_item(ui, icon::FILE_PDF, "Export PDF", false).clicked() {
                        self.cmds.push(Cmd::ExportPdf);
                    }
                    if nav_item(ui, icon::CLOUD_ARROW_UP, "Upload to Intune", false).clicked() {
                        self.cmds.push(Cmd::Backstage(None));
                        self.cmds.push(Cmd::IntuneUpload);
                    }
                    if nav_item(ui, icon::X_CIRCLE, "Close project", false).clicked() {
                        self.cmds.push(Cmd::CloseProject);
                    }
                }
                ui.with_layout(egui::Layout::bottom_up(egui::Align::LEFT), |ui| {
                    if nav_item(ui, icon::SIGN_OUT, "Exit", false).clicked() {
                        self.cmds.push(Cmd::Quit);
                    }
                    let dark_label = if self.prefs.dark_mode { "Light mode" } else { "Dark mode" };
                    let dark_icon = if self.prefs.dark_mode { icon::SUN } else { icon::MOON };
                    if nav_item(ui, dark_icon, dark_label, false).clicked() {
                        self.cmds.push(Cmd::ToggleDark);
                    }
                    if nav_item(ui, icon::INFO, "About", page == BackstagePage::About).clicked() {
                        self.cmds.push(Cmd::Backstage(Some(BackstagePage::About)));
                    }
                });
            });

        egui::CentralPanel::default()
            .frame(egui::Frame::new().fill(p.window_bg).inner_margin(egui::Margin { left: 48, right: 48, top: 36, bottom: 24 }))
            .show(ui, |ui| {
                egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
                    ui.set_max_width(1100.0);
                    match page {
                        BackstagePage::Home => self.bs_home(ui),
                        BackstagePage::New => self.bs_new(ui),
                        BackstagePage::Open => self.bs_open(ui),
                        BackstagePage::Info => self.bs_info(ui),
                        BackstagePage::About => self.bs_about(ui),
                    }
                });
            });
    }

    fn greeting(&self) -> &'static str {
        use chrono::Timelike;
        match chrono::Local::now().hour() {
            5..=11 => "Good morning",
            12..=17 => "Good afternoon",
            _ => "Good evening",
        }
    }

    fn bs_home(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        ui.label(egui::RichText::new(self.greeting()).font(semibold(28.0)));
        muted(ui, "Build PSAppDeployToolkit packages visually — no hand-written deployment scripts.");
        ui.add_space(22.0);
        ui.label(egui::RichText::new("New").font(semibold(16.0)));
        ui.add_space(8.0);
        self.template_cards(ui, false);
        ui.add_space(28.0);
        ui.horizontal(|ui| {
            ui.label(egui::RichText::new("Recent").font(semibold(16.0)));
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if ui.link(it(format!("{}  Browse…", icon::FOLDER_OPEN))).clicked() {
                    self.cmds.push(Cmd::BrowseOpenProject);
                }
            });
        });
        ui.add_space(6.0);
        self.recent_list(ui);
        let _ = p;
    }

    fn bs_new(&mut self, ui: &mut egui::Ui) {
        ui.label(egui::RichText::new("New").font(semibold(28.0)));
        muted(ui, "Every project is a self-contained folder with the PSADT framework and an Invoke-AppDeployToolkit.db.");
        ui.add_space(20.0);
        self.template_cards(ui, true);
    }

    fn template_cards(&mut self, ui: &mut egui::Ui, large: bool) {
        let cards: [(&str, &str, &str, Cmd); 5] = [
            (icon::FILE_PLUS, "Blank project", "Start with an empty install, uninstall and repair sequence.", Cmd::NewProject(NewKind::Blank)),
            (icon::DISC, "Import MSI", "Reads product name, version and code — creates install & uninstall steps.", Cmd::NewProject(NewKind::Msi(String::new()))),
            (icon::APP_WINDOW, "Import EXE", "Copies the installer into Files and adds a Start-ADTProcess step.", Cmd::NewProject(NewKind::Exe(String::new()))),
            (icon::DOWNLOAD_SIMPLE, "Import from WinGet", "Search the WinGet catalog, download the installer and build the sequence.", Cmd::WinGetImport),
            (icon::PACKAGE, "Demo project", "Open the bundled 7-Zip example to explore the editor.", Cmd::OpenDemoProject),
        ];
        let h = 152.0_f32; let _ = large;
        ui.horizontal_wrapped(|ui| {
            ui.spacing_mut().item_spacing = Vec2::new(14.0, 14.0);
            for (ic, title, desc, cmd) in cards {
                if card(ui, ic, title, desc, Vec2::new(204.0, h)).clicked() {
                    self.cmds.push(cmd);
                }
            }
        });
    }

    fn recent_list(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        if self.prefs.recent.is_empty() {
            egui::Frame::new()
                .fill(p.panel_bg)
                .stroke(Stroke::new(1.0, p.border))
                .corner_radius(CornerRadius::same(10))
                .inner_margin(egui::Margin::same(18))
                .show(ui, |ui| {
                    ui.set_width(ui.available_width());
                    muted(ui, "Projects you open will appear here.");
                });
            return;
        }
        egui::Frame::new()
            .fill(p.panel_bg)
            .stroke(Stroke::new(1.0, p.border))
            .corner_radius(CornerRadius::same(10))
            .inner_margin(egui::Margin::same(6))
            .show(ui, |ui| {
                ui.set_width(ui.available_width());
                for db in self.prefs.recent.clone() {
                    let exists = db.is_file();
                    let folder = db.parent().map(|f| f.display().to_string()).unwrap_or_default();
                    let name = db
                        .parent()
                        .and_then(|f| f.file_name())
                        .map(|n| n.to_string_lossy().to_string())
                        .unwrap_or_else(|| db.display().to_string());
                    let (rect, resp) = ui.allocate_exact_size(Vec2::new(ui.available_width(), 48.0), Sense::click());
                    if resp.hovered() {
                        ui.painter().rect_filled(rect, CornerRadius::same(8), p.hover);
                    }
                    ui.painter().text(
                        rect.left_center() + Vec2::new(22.0, 0.0),
                        Align2::CENTER_CENTER,
                        icon::PACKAGE,
                        crate::theme::icon_font(22.0),
                        if exists { p.accent } else { p.text_faint },
                    );
                    ui.painter().text(
                        rect.left_top() + Vec2::new(46.0, 8.0),
                        Align2::LEFT_TOP,
                        &name,
                        semibold(14.0),
                        if exists { p.text } else { p.text_faint },
                    );
                    let sub = if exists { folder.clone() } else { format!("{folder}  (not found)") };
                    ui.painter().text(
                        rect.left_top() + Vec2::new(46.0, 27.0),
                        Align2::LEFT_TOP,
                        sub,
                        FontId::proportional(12.0),
                        p.text_muted,
                    );
                    let x_rect = egui::Rect::from_center_size(rect.right_center() - Vec2::new(24.0, 0.0), Vec2::splat(26.0));
                    let x_resp = ui.interact(x_rect, egui::Id::new(("forget", &db)), Sense::click());
                    if resp.hovered() || x_resp.hovered() {
                        if x_resp.hovered() {
                            ui.painter().rect_filled(x_rect, CornerRadius::same(6), p.pressed);
                        }
                        ui.painter().text(x_rect.center(), Align2::CENTER_CENTER, icon::X, crate::theme::icon_font(14.0), p.text_muted);
                    }
                    if x_resp.on_hover_text("Remove from list").clicked() {
                        self.cmds.push(Cmd::ForgetRecent(db.clone()));
                    } else if resp.clicked() && exists {
                        self.cmds.push(Cmd::OpenProject(db.clone()));
                    }
                }
            });
    }

    fn bs_open(&mut self, ui: &mut egui::Ui) {
        ui.label(egui::RichText::new("Open").font(semibold(28.0)));
        muted(ui, format!("Choose a project database ({DATABASE_FILE}) — projects from the PureBasic version open unchanged.").as_str());
        ui.add_space(16.0);
        ui.horizontal(|ui| {
            if primary_button(ui, &format!("{}  Browse…", icon::FOLDER_OPEN)).clicked() {
                self.cmds.push(Cmd::BrowseOpenProject);
            }
            if secondary_button(ui, &format!("{}  Demo project", icon::PACKAGE)).clicked() {
                self.cmds.push(Cmd::OpenDemoProject);
            }
        });
        ui.add_space(22.0);
        ui.label(egui::RichText::new("Recent").font(semibold(16.0)));
        ui.add_space(6.0);
        self.recent_list(ui);
    }

    fn bs_info(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        let Some(project) = &self.project else { return };
        let s = project.settings();
        let folder = project.folder.clone();
        ui.label(egui::RichText::new("Info").font(semibold(28.0)));
        muted(ui, &folder.display().to_string());
        ui.add_space(18.0);
        ui.horizontal_top(|ui| {
            egui::Frame::new()
                .fill(p.panel_bg)
                .stroke(Stroke::new(1.0, p.border))
                .corner_radius(CornerRadius::same(10))
                .inner_margin(egui::Margin::same(18))
                .show(ui, |ui| ui.vertical(|ui| {
                    ui.set_width(420.0);
                    ui.label(egui::RichText::new("Application").font(semibold(15.0)));
                    ui.add_space(6.0);
                    egui::Grid::new("info-grid").num_columns(2).spacing(Vec2::new(18.0, 8.0)).show(ui, |ui| {
                        for (k, v) in [
                            ("Project", &s.project_name),
                            ("Name", &s.app_name),
                            ("Version", &s.app_version),
                            ("Vendor", &s.app_vendor),
                            ("Architecture", &s.app_architecture),
                            ("Language", &s.app_language),
                            ("Author", &s.app_author),
                        ] {
                            ui.label(egui::RichText::new(k).color(p.text_muted));
                            ui.label(if v.is_empty() { "—" } else { v.as_str() });
                            ui.end_row();
                        }
                    });
                    ui.add_space(10.0);
                    if primary_button(ui, &format!("{}  Edit settings", icon::GEAR_SIX)).clicked() {
                        self.cmds.push(Cmd::ProjectSettings);
                    }
                }));
            ui.add_space(14.0);
            egui::Frame::new()
                .fill(p.panel_bg)
                .stroke(Stroke::new(1.0, p.border))
                .corner_radius(CornerRadius::same(10))
                .inner_margin(egui::Margin::same(18))
                .show(ui, |ui| ui.vertical(|ui| {
                    ui.set_width(320.0);
                    ui.label(egui::RichText::new("Sequences").font(semibold(15.0)));
                    ui.add_space(6.0);
                    for (i, label) in ["Installation", "Uninstall", "Repair"].iter().enumerate() {
                        ui.horizontal(|ui| {
                            ui.label(*label);
                            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                let n = self.counts[i];
                                ui.label(it(format!("{n} step{}", if n == 1 { "" } else { "s" })).color(p.text_muted));
                            });
                        });
                    }
                    ui.add_space(10.0);
                    let gen = folder.join(crate::project::DEPLOYMENT_FILE);
                    let when = std::fs::metadata(&gen)
                        .and_then(|m| m.modified())
                        .ok()
                        .map(|t| chrono::DateTime::<chrono::Local>::from(t).format("%Y-%m-%d %H:%M").to_string());
                    ui.label(egui::RichText::new("Deployment file").font(semibold(13.0)));
                    muted(ui, &when.map(|w| format!("Last generated {w}")).unwrap_or_else(|| "Not generated yet".into()));
                }));
        });
    }

    fn bs_about(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        ui.horizontal(|ui| {
            let (r, _) = ui.allocate_exact_size(Vec2::splat(72.0), Sense::hover());
            ui.painter().rect_filled(r, CornerRadius::same(18), p.accent);
            ui.painter().text(r.center(), Align2::CENTER_CENTER, icon::PACKAGE, crate::theme::icon_font(42.0), p.on_accent);
            ui.add_space(10.0);
            ui.vertical(|ui| {
                ui.label(egui::RichText::new("Deployment Editor").font(semibold(28.0)));
                muted(ui, &format!(
                    "Version {} ({}) · Rust edition · by TUGI.CH",
                    env!("CARGO_PKG_VERSION"),
                    match std::env::consts::ARCH { "aarch64" => "ARM64", "x86_64" => "x64", a => a }
                ));
                muted(ui, &format!("Graphics: {}", self.graphics));
            });
        });
        ui.add_space(18.0);
        ui.label("Visual software packaging editor for PSAppDeployToolkit (PSADT). Build installation, uninstall and repair logic visually; the editor generates a production-ready Invoke-AppDeployToolkit.ps1.");
        ui.add_space(14.0);
        ui.horizontal_wrapped(|ui| {
            for (label, url) in [
                ("GitHub", "https://github.com/tugich/DeploymentEditor"),
                ("www.tugi.ch", "https://www.tugi.ch"),
                ("PSAppDeployToolkit", "https://psappdeploytoolkit.com"),
                ("Donate", "https://www.paypal.com/donate/?hosted_button_id=PXABL8ESQQ4F8"),
            ] {
                if secondary_button(ui, label).clicked() {
                    self.cmds.push(Cmd::OpenUrl(url.into()));
                }
            }
        });
        ui.add_space(18.0);
        muted(ui, "PSAppDeployToolkit is licensed under the GNU LGPLv3. Icons: Phosphor (MIT). UI: egui (MIT/Apache-2.0).");
        ui.add_space(6.0);
        muted(ui, &format!("Resources: {}", self.paths.root.display()));
        if let Some(e) = &self.catalog_error {
            ui.add_space(8.0);
            ui.label(egui::RichText::new(e).color(p.danger));
        }
    }
}

fn nav_item(ui: &mut egui::Ui, ic: &str, label: &str, selected: bool) -> egui::Response {
    let p = crate::widgets::palette(ui);
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(ui.available_width(), 38.0), Sense::click());
    let overlay = if p.dark { egui::Color32::from_black_alpha(50) } else { egui::Color32::from_white_alpha(38) };
    if selected {
        ui.painter().rect_filled(rect, CornerRadius::same(8), overlay);
        ui.painter().rect_filled(
            egui::Rect::from_min_size(rect.left_top() + Vec2::new(0.0, 10.0), Vec2::new(3.0, 18.0)),
            CornerRadius::same(2),
            p.on_accent,
        );
    } else if resp.hovered() {
        ui.painter().rect_filled(rect, CornerRadius::same(8), overlay.gamma_multiply(0.6));
    }
    ui.painter().text(rect.left_center() + Vec2::new(22.0, 0.0), Align2::CENTER_CENTER, ic, crate::theme::icon_font(18.0), p.on_accent);
    ui.painter().text(
        rect.left_center() + Vec2::new(42.0, 0.0),
        Align2::LEFT_CENTER,
        label,
        if selected { semibold(14.0) } else { FontId::proportional(14.0) },
        p.on_accent,
    );
    resp.on_hover_cursor(egui::CursorIcon::PointingHand)
}

fn card(ui: &mut egui::Ui, ic: &str, title: &str, desc: &str, size: Vec2) -> egui::Response {
    let p = crate::widgets::palette(ui);
    let galley = ui.painter().layout(desc.to_string(), FontId::proportional(12.0), p.text_muted, size.x - 32.0);
    let height = size.y.max(88.0 + galley.size().y + 16.0);
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(size.x, height), Sense::click());
    let hovered = resp.hovered();
    ui.painter().rect(
        rect,
        CornerRadius::same(12),
        if hovered { p.card_hover } else { p.card_bg },
        Stroke::new(if hovered { 1.5 } else { 1.0 }, if hovered { p.accent } else { p.border }),
        StrokeKind::Inside,
    );
    let badge = egui::Rect::from_min_size(rect.left_top() + Vec2::new(16.0, 16.0), Vec2::splat(40.0));
    ui.painter().rect_filled(badge, CornerRadius::same(10), p.accent_soft);
    ui.painter().text(badge.center(), Align2::CENTER_CENTER, ic, crate::theme::icon_font(22.0), p.accent);
    ui.painter().text(rect.left_top() + Vec2::new(16.0, 66.0), Align2::LEFT_TOP, title, semibold(14.5), p.text);
    ui.painter().galley(rect.left_top() + Vec2::new(16.0, 88.0), galley, p.text_muted);
    resp.on_hover_cursor(egui::CursorIcon::PointingHand)
}
