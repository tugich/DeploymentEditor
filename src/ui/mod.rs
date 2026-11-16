//! Main window layout: ribbon, panels, status bar, toasts, shortcuts.

mod backstage;
pub mod condition_builder;
mod detection_dialog;
pub mod intune_upload;
mod dialogs;
mod inspector;
mod library;
mod ribbon;
mod sequence;

use eframe::egui::{self, Align2, Color32, CornerRadius, FontId, Key, Modifiers, Stroke, Vec2};
use egui_phosphor::regular as icon;

use crate::app::{Cmd, DeploymentEditorApp, ToastKind};
use crate::project::DeploymentType;
use crate::theme::semibold;
use crate::widgets::it;

/// Payload for dragging an action card.
#[derive(Clone, Copy, Debug)]
pub struct DragAction(pub i64);
/// Payload for dragging a command from the library.
#[derive(Clone, Debug)]
pub struct DragCommand(pub String);

impl eframe::App for DeploymentEditorApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        let ctx = ui.ctx().clone();
        self.poll_jobs();
        self.handle_input(&ctx);

        let title = self.window_title();
        if title != self.window_title {
            ctx.send_viewport_cmd(egui::ViewportCommand::Title(title.clone()));
            self.window_title = title;
        }

        let p = self.palette;
        let show_workspace = self.backstage.is_none() && self.project.is_some();

        if show_workspace {
            self.tab_row(ui);
            if !self.prefs.ribbon_collapsed {
                self.ribbon(ui);
            }
        }
        self.status_bar(ui);

        if !show_workspace {
            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(p.window_bg))
                .show(ui, |ui| self.backstage(ui));
        } else {
            if self.prefs.show_library {
                egui::Panel::left("library")
                    .resizable(true)
                    .default_size(290.0)
                    .size_range(220.0..=460.0)
                    .frame(
                        egui::Frame::new()
                            .fill(p.panel_bg)
                            .inner_margin(egui::Margin::symmetric(12, 12))
                            .stroke(Stroke::new(1.0, p.border)),
                    )
                    .show(ui, |ui| self.library(ui));
            }
            if self.prefs.show_inspector {
                egui::Panel::right("inspector")
                    .resizable(true)
                    .default_size(380.0)
                    .size_range(300.0..=640.0)
                    .frame(
                        egui::Frame::new()
                            .fill(p.panel_bg)
                            .inner_margin(egui::Margin::symmetric(16, 12))
                            .stroke(Stroke::new(1.0, p.border)),
                    )
                    .show(ui, |ui| self.inspector(ui));
            }
            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(p.canvas_bg).inner_margin(egui::Margin::symmetric(20, 14)))
                .show(ui, |ui| self.sequence(ui));
        }

        self.dialogs(&ctx);
        if !self.upload_hidden {
            self.upload_window(&ctx);
        }
        self.preview_window(&ctx);
        self.script_editor_window(&ctx);
        self.log_viewer_window(&ctx);
        self.file_hover_overlay(&ctx);
        self.toasts_overlay(&ctx);

        self.run_queued();
    }

    fn save(&mut self, storage: &mut dyn eframe::Storage) {
        self.save_prefs(storage);
    }

    fn on_exit(&mut self, _gl: Option<&eframe::glow::Context>) {
        self.commit_draft();
    }

    fn clear_color(&self, _visuals: &egui::Visuals) -> [f32; 4] {
        self.palette.window_bg.to_normalized_gamma_f32()
    }
}

impl DeploymentEditorApp {
    fn handle_input(&mut self, ctx: &egui::Context) {
        let dropped: Vec<std::path::PathBuf> =
            ctx.input(|i| i.raw.dropped_files.iter().map(|f| f.path().to_path_buf()).filter(|p| !p.as_os_str().is_empty()).collect());
        if !dropped.is_empty() {
            self.cmds.push(Cmd::DropFiles(dropped));
        }

        let typing = ctx.egui_wants_keyboard_input();
        let has_project = self.project.is_some();
        let modal_open = self.dialog.is_some();
        let mut push = |c: Cmd| self.cmds.push(c);

        ctx.input_mut(|i| {
            if i.consume_key(Modifiers::COMMAND, Key::S) {
                push(Cmd::Save);
            }
            if i.consume_key(Modifiers::COMMAND, Key::O) {
                push(Cmd::BrowseOpenProject);
            }
            if i.consume_key(Modifiers::COMMAND, Key::N) {
                push(Cmd::NewProject(crate::app::NewKind::Blank));
            }
            if modal_open || !has_project {
                return;
            }
            if i.consume_key(Modifiers::NONE, Key::F5) {
                push(Cmd::Run(DeploymentType::Installation));
            }
            if i.consume_key(Modifiers::COMMAND, Key::G) {
                push(Cmd::Generate);
            }
            if i.consume_key(Modifiers::COMMAND, Key::P) {
                push(Cmd::Preview);
            }
            if i.consume_key(Modifiers::COMMAND, Key::D) {
                push(Cmd::Duplicate);
            }
            if i.consume_key(Modifiers::COMMAND, Key::E) {
                push(Cmd::ToggleDisable(None));
            }
            if i.consume_key(Modifiers::ALT, Key::ArrowUp) {
                push(Cmd::Move(-1));
            }
            if i.consume_key(Modifiers::ALT, Key::ArrowDown) {
                push(Cmd::Move(1));
            }
            if i.consume_key(Modifiers::NONE, Key::F7) {
                push(Cmd::ShowProblems);
            }
            if i.consume_key(Modifiers::NONE, Key::F1) {
                push(Cmd::CommandHelp);
            }
            // Text fields keep their own undo; the project history applies elsewhere.
            if !typing {
                if i.consume_key(Modifiers::COMMAND | Modifiers::SHIFT, Key::Z) || i.consume_key(Modifiers::COMMAND, Key::Y) {
                    push(Cmd::Redo);
                } else if i.consume_key(Modifiers::COMMAND, Key::Z) {
                    push(Cmd::Undo);
                }
            }
            if !typing && i.consume_key(Modifiers::NONE, Key::Delete) {
                push(Cmd::RemoveSelected);
            }
        });

        // Arrow key navigation in the sequence when not typing.
        if !typing && !modal_open && has_project && !self.actions.is_empty() {
            let (up, down) = ctx.input(|i| {
                (
                    i.key_pressed(Key::ArrowUp) && i.modifiers.is_none(),
                    i.key_pressed(Key::ArrowDown) && i.modifiers.is_none(),
                )
            });
            if up || down {
                let pos = self.selected.and_then(|id| self.actions.iter().position(|a| a.id == id));
                let next = match pos {
                    None => 0,
                    Some(p) if up => p.saturating_sub(1),
                    Some(p) => (p + 1).min(self.actions.len() - 1),
                };
                let id = self.actions[next].id;
                self.cmds.push(Cmd::Select(Some(id)));
                self.scroll_to_selected = true;
            }
        }
    }

    fn tab_row(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        egui::Panel::top("tabs")
            .exact_size(40.0)
            .show_separator_line(false)
            .frame(egui::Frame::new().fill(p.ribbon_bg).inner_margin(egui::Margin { left: 8, right: 10, top: 6, bottom: 0 }))
            .show(ui, |ui| {
                ui.horizontal(|ui| {
                    ui.spacing_mut().item_spacing.x = 2.0;
                    // File button (opens backstage)
                    let (rect, resp) = ui.allocate_exact_size(Vec2::new(56.0, 30.0), egui::Sense::click());
                    let fill = if resp.hovered() { p.accent_hover } else { p.accent };
                    ui.painter().rect_filled(rect, CornerRadius::same(6), fill);
                    ui.painter().text(rect.center(), Align2::CENTER_CENTER, "File", semibold(13.5), p.on_accent);
                    if resp.on_hover_cursor(egui::CursorIcon::PointingHand).clicked() {
                        self.cmds.push(Cmd::Backstage(Some(crate::app::BackstagePage::Home)));
                    }
                    ui.add_space(4.0);
                    // Quick access toolbar
                    let undo_tip = match self.undo_label() {
                        Some(l) => format!("Undo {l} (Ctrl+Z)"),
                        None => "Nothing to undo".into(),
                    };
                    let redo_tip = match self.redo_label() {
                        Some(l) => format!("Redo {l} (Ctrl+Y)"),
                        None => "Nothing to redo".into(),
                    };
                    let can_undo = !self.undo_stack.is_empty() || self.dirty;
                    let can_redo = !self.redo_stack.is_empty();
                    if ui.add_enabled_ui(can_undo, |ui| crate::widgets::icon_button(ui, icon::ARROW_U_UP_LEFT, &undo_tip)).inner.clicked() {
                        self.cmds.push(Cmd::Undo);
                    }
                    if ui.add_enabled_ui(can_redo, |ui| crate::widgets::icon_button(ui, icon::ARROW_U_UP_RIGHT, &redo_tip)).inner.clicked() {
                        self.cmds.push(Cmd::Redo);
                    }
                    ui.add_space(8.0);

                    for tab in crate::app::RibbonTab::ALL {
                        let selected = self.ribbon_tab == tab;
                        let font = if selected { semibold(13.5) } else { FontId::proportional(13.5) };
                        let galley = ui.painter().layout_no_wrap(tab.label().to_string(), font.clone(), p.text);
                        let w = galley.size().x + 24.0;
                        let (rect, resp) = ui.allocate_exact_size(Vec2::new(w, 34.0), egui::Sense::click());
                        if resp.hovered() && !selected {
                            ui.painter().rect_filled(rect.shrink2(Vec2::new(0.0, 3.0)), CornerRadius::same(6), p.hover);
                        }
                        let color = if selected { p.accent_text } else { p.text };
                        ui.painter().text(rect.center() - Vec2::new(0.0, 2.0), Align2::CENTER_CENTER, tab.label(), font, color);
                        if selected {
                            let bar = egui::Rect::from_center_size(
                                egui::pos2(rect.center().x, rect.bottom() - 3.0),
                                Vec2::new(w - 20.0, 3.0),
                            );
                            ui.painter().rect_filled(bar, CornerRadius::same(2), p.accent);
                        }
                        if resp.clicked() {
                            if selected && self.prefs.ribbon_collapsed {
                                self.prefs.ribbon_collapsed = false;
                            }
                            self.ribbon_tab = tab;
                        }
                        if resp.double_clicked() {
                            self.prefs.ribbon_collapsed = !self.prefs.ribbon_collapsed;
                        }
                    }

                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        let chevron = if self.prefs.ribbon_collapsed { icon::CARET_DOWN } else { icon::CARET_UP };
                        let tip = if self.prefs.ribbon_collapsed { "Expand the ribbon" } else { "Collapse the ribbon" };
                        if crate::widgets::icon_button(ui, chevron, tip).clicked() {
                            self.prefs.ribbon_collapsed = !self.prefs.ribbon_collapsed;
                        }
                        let theme_icon = if self.prefs.dark_mode { icon::SUN } else { icon::MOON };
                        if crate::widgets::icon_button(ui, theme_icon, "Toggle dark mode").clicked() {
                            self.cmds.push(Cmd::ToggleDark);
                        }
                        ui.add_space(6.0);
                        self.quick_search(ui);
                        ui.add_space(10.0);
                        // project chip
                        let text = format!("{}  {}", icon::PACKAGE, self.project_name);
                        let r = ui.add(
                            egui::Button::new(it(text).font(semibold(13.0)).color(p.text))
                                .frame(false),
                        );
                        if r.on_hover_text("Project settings").clicked() {
                            self.cmds.push(Cmd::ProjectSettings);
                        }
                        if self.dirty {
                            ui.label(it(icon::CIRCLE).size(9.0).color(p.warning))
                                .on_hover_text("Unsaved changes (Ctrl+S)");
                        }
                    });
                });
            });
    }

    /// Office-style "Tell me" search: type a command, press Enter to add it.
    fn quick_search(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        let id = egui::Id::new("quick-search");
        let resp = ui.add(
            egui::TextEdit::singleline(&mut self.quick_search)
                .id(id)
                .desired_width(260.0)
                .hint_text(it(format!("{}  Search commands to add…  (Ctrl+K)", icon::MAGNIFYING_GLASS)))
                .margin(egui::Margin::symmetric(10, 5)),
        );
        if ui.input_mut(|i| i.consume_key(Modifiers::COMMAND, Key::K)) {
            resp.request_focus();
        }
        let q = self.quick_search.trim().to_lowercase();
        if q.is_empty() || !(resp.has_focus() || resp.lost_focus() || self.quick_popup_hovered) {
            self.quick_popup_hovered = false;
            return;
        }
        let matches: Vec<(String, String)> = self
            .catalog
            .visible_commands()
            .filter(|c| c.command.to_lowercase().contains(&q) || c.name.to_lowercase().contains(&q))
            .take(8)
            .map(|c| (c.command.clone(), c.description.clone()))
            .collect();
        let enter = resp.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter));
        if enter {
            if let Some((c, _)) = matches.first() {
                self.cmds.push(Cmd::AddCommand(c.clone()));
            }
            self.quick_search.clear();
            return;
        }
        if ui.input(|i| i.key_pressed(Key::Escape)) {
            self.quick_search.clear();
            return;
        }
        let pos = resp.rect.right_bottom() + Vec2::new(-436.0, 4.0);
        let area = egui::Area::new(id.with("popup"))
            .order(egui::Order::Foreground)
            .fixed_pos(pos)
            .show(ui.ctx(), |ui| {
                egui::Frame::popup(ui.style()).fill(p.panel_bg).show(ui, |ui| {
                    ui.set_width(420.0);
                    if matches.is_empty() {
                        crate::widgets::muted(ui, "No matching PSADT command");
                    }
                    for (i, (c, d)) in matches.iter().enumerate() {
                        let mut short: String = d.chars().take(70).collect();
                        if d.chars().count() > 70 {
                            short.push('…');
                        }
                        let r = ui.add(
                            egui::Button::new(
                                it(format!("{}  {c}", icon::PLUS_CIRCLE)).font(semibold(13.0)),
                            )
                            .right_text(egui::RichText::new(if i == 0 { "Enter" } else { "" }).size(11.0).color(p.text_faint))
                            .frame(false)
                            .min_size(Vec2::new(410.0, 24.0)),
                        );
                        ui.label(egui::RichText::new(short).size(11.5).color(p.text_muted));
                        if r.clicked() {
                            self.cmds.push(Cmd::AddCommand(c.clone()));
                            self.quick_search.clear();
                        }
                    }
                });
            });
        self.quick_popup_hovered = area.response.contains_pointer();
    }

    fn status_bar(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        let bg = if self.project.is_some() && self.backstage.is_none() { p.ribbon_bg } else { p.window_bg };
        egui::Panel::bottom("status")
            .exact_size(28.0)
            .frame(
                egui::Frame::new()
                    .fill(bg)
                    .inner_margin(egui::Margin::symmetric(12, 4))
                    .stroke(Stroke::new(1.0, p.border)),
            )
            .show(ui, |ui| {
                ui.horizontal_centered(|ui| {
                    ui.spacing_mut().item_spacing.x = 14.0;
                    if let Some(b) = &self.busy {
                        ui.add(egui::Spinner::new().size(13.0));
                        ui.label(egui::RichText::new(b).size(12.0).color(p.text_muted));
                    } else {
                        ui.label(
                            it(format!("{}  {}", icon::INFO, self.status)).size(12.0).color(p.text_muted),
                        );
                    }
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        ui.label(egui::RichText::new("TUGI.CH").size(11.5).color(p.text_faint));
                        if self.project.is_some() {
                            let errors = self.error_count();
                            let warnings = crate::validate::count(&self.issues, crate::validate::Severity::Warning);
                            let (text, color) = if errors > 0 {
                                (format!("{}  {errors} error{}", icon::WARNING_CIRCLE, if errors == 1 { "" } else { "s" }), p.danger)
                            } else if warnings > 0 {
                                (format!("{}  {warnings} warning{}", icon::WARNING, if warnings == 1 { "" } else { "s" }), p.warning)
                            } else {
                                (format!("{}  No problems", icon::CHECK_CIRCLE), p.success)
                            };
                            if ui
                                .add(egui::Label::new(it(text).size(12.0).color(color)).sense(egui::Sense::click()))
                                .on_hover_text("Show problems")
                                .clicked()
                            {
                                self.cmds.push(Cmd::ShowProblems);
                            }
                        }
                        if let Some(proj) = &self.project {
                            let path = proj.folder.display().to_string();
                            if ui
                                .add(
                                    egui::Label::new(
                                        it(format!("{}  {path}", icon::FOLDER)).size(12.0).color(p.text_muted),
                                    )
                                    .sense(egui::Sense::click()),
                                )
                                .on_hover_text("Open the project folder")
                                .clicked()
                            {
                                self.cmds.push(Cmd::OpenFolder(crate::app::FolderKind::Project));
                            }
                            if let Some(t) = &self.last_generated {
                                ui.label(
                                    it(format!("{}  Generated {t}", icon::CHECK_CIRCLE))
                                        .size(12.0)
                                        .color(p.success),
                                );
                            }
                            if self.prefs.silent_mode {
                                ui.label(it(format!("{}  Silent", icon::SPEAKER_SLASH)).size(12.0).color(p.text_muted));
                            }
                            ui.label(
                                it(format!(
                                    "Install {}  ·  Uninstall {}  ·  Repair {}",
                                    self.counts[0], self.counts[1], self.counts[2]
                                ))
                                .size(12.0)
                                .color(p.text_muted),
                            );
                        }
                    });
                });
            });
    }

    fn file_hover_overlay(&mut self, ctx: &egui::Context) {
        let hovering = ctx.input(|i| !i.raw.hovered_files.is_empty());
        if !hovering {
            return;
        }
        let p = self.palette;
        let rect = ctx.content_rect();
        let painter = ctx.layer_painter(egui::LayerId::new(egui::Order::Foreground, egui::Id::new("drop")));
        painter.rect_filled(rect, CornerRadius::ZERO, Color32::from_black_alpha(90));
        let card = egui::Rect::from_center_size(rect.center(), Vec2::new(420.0, 140.0));
        painter.rect(card, CornerRadius::same(14), p.panel_bg, Stroke::new(2.0, p.accent), egui::StrokeKind::Inside);
        let msg = if self.project.is_some() {
            "Drop files to add them as actions\nEXE · MSI · LNK · PS1 — or a project .db to open it"
        } else {
            "Drop a project database (.db) to open it"
        };
        painter.text(card.center() - Vec2::new(0.0, 26.0), Align2::CENTER_CENTER, icon::DOWNLOAD_SIMPLE, crate::theme::icon_font(30.0), p.accent);
        painter.text(card.center() + Vec2::new(0.0, 22.0), Align2::CENTER_CENTER, msg, FontId::proportional(14.0), p.text);
    }

    fn toasts_overlay(&mut self, ctx: &egui::Context) {
        let p = self.palette;
        self.toasts.retain(|t| t.created.elapsed().as_secs_f32() < 4.5);
        if self.toasts.is_empty() {
            return;
        }
        ctx.request_repaint_after(std::time::Duration::from_millis(100));
        let rect = ctx.content_rect();
        egui::Area::new(egui::Id::new("toasts"))
            .order(egui::Order::Tooltip)
            .anchor(Align2::CENTER_BOTTOM, Vec2::new(0.0, -44.0))
            .interactable(false)
            .show(ctx, |ui| {
                ui.set_max_width((rect.width() * 0.4).clamp(280.0, 520.0));
                for t in &self.toasts {
                    let age = t.created.elapsed().as_secs_f32();
                    let alpha = if age > 4.0 { (4.5 - age) / 0.5 } else { (age / 0.15).min(1.0) };
                    let (ic, col) = match t.kind {
                        ToastKind::Info => (icon::INFO, p.info),
                        ToastKind::Success => (icon::CHECK_CIRCLE, p.success),
                        ToastKind::Error => (icon::WARNING_CIRCLE, p.danger),
                    };
                    ui.scope(|ui| {
                        ui.set_opacity(alpha.clamp(0.0, 1.0));
                        egui::Frame::new()
                            .fill(p.panel_bg)
                            .stroke(Stroke::new(1.0, p.border))
                            .corner_radius(CornerRadius::same(10))
                            .shadow(egui::epaint::Shadow { offset: [0, 4], blur: 18, spread: 0, color: Color32::from_black_alpha(40) })
                            .inner_margin(egui::Margin::symmetric(14, 10))
                            .show(ui, |ui| {
                                ui.horizontal(|ui| {
                                    ui.label(it(ic).size(18.0).color(col));
                                    ui.add(egui::Label::new(egui::RichText::new(&t.text).size(13.0)).wrap());
                                });
                            });
                    });
                    ui.add_space(6.0);
                }
            });
    }
}
