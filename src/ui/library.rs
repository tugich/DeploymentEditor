//! Left pane: PSADT command library (searchable, draggable) and snippets.

use eframe::egui::{self, Align2, CornerRadius, FontId, Sense, Vec2};
use egui_phosphor::regular as icon;

use super::DragCommand;
use crate::app::{Cmd, DeploymentEditorApp, FolderKind};
use crate::theme::semibold;
use crate::widgets::it;

/// Commands most packagers reach for first.
const COMMON: &[&str] = &[
    "Start-ADTProcess",
    "Start-ADTMsiProcess",
    "Uninstall-ADTApplication",
    "Show-ADTInstallationWelcome",
    "Show-ADTInstallationProgress",
    "Show-ADTInstallationPrompt",
    "Copy-ADTFile",
    "Remove-ADTFile",
    "New-ADTShortcut",
    "Set-ADTRegistryKey",
    "Remove-ADTRegistryKey",
    "Get-ADTApplication",
];

pub fn verb_icon(command: &str) -> &'static str {
    let verb = command.split('-').next().unwrap_or("");
    match verb {
        "Get" => icon::MAGNIFYING_GLASS,
        "Set" | "Update" => icon::PENCIL_SIMPLE,
        "New" | "Add" => icon::PLUS_SQUARE,
        "Remove" | "Clear" | "Uninstall" | "Unregister" => icon::TRASH_SIMPLE,
        "Start" | "Invoke" => icon::PLAY_CIRCLE,
        "Stop" | "Block" | "Disable" => icon::STOP_CIRCLE,
        "Show" => icon::APP_WINDOW,
        "Close" => icon::X_SQUARE,
        "Copy" => icon::COPY,
        "Test" => icon::CHECK_SQUARE,
        "Install" | "Register" | "Enable" => icon::DOWNLOAD_SIMPLE,
        "Convert" | "ConvertTo" => icon::ARROWS_LEFT_RIGHT,
        "Mount" | "Dismount" => icon::HARD_DRIVES,
        "Write" | "Out" | "Export" => icon::FILE_TEXT,
        "Send" => icon::PAPER_PLANE_TILT,
        _ => icon::TERMINAL,
    }
}

impl DeploymentEditorApp {
    pub(super) fn library(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;

        // Snippets at the bottom
        egui::Panel::bottom("snippets")
            .resizable(false)
            .show_separator_line(true)
            .frame(egui::Frame::new().inner_margin(egui::Margin { left: 0, right: 0, top: 8, bottom: 0 }))
            .show(ui, |ui| {
                ui.horizontal(|ui| {
                    ui.label(egui::RichText::new("Snippets").font(semibold(13.5)));
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if crate::widgets::icon_button(ui, icon::FOLDER_OPEN, "Open Snippets folder").clicked() {
                            self.cmds.push(Cmd::OpenFolder(FolderKind::Snippets));
                        }
                    });
                });
                egui::ScrollArea::vertical().id_salt("snip").max_height(110.0).show(ui, |ui| {
                    if self.snippets.is_empty() {
                        crate::widgets::muted(ui, "Put .ps1 files into the Snippets folder.");
                    }
                    for (i, s) in self.snippets.iter().enumerate() {
                        let r = list_row(ui, icon::FILE_CODE, &s.file_name, false, p.script);
                        if r.on_hover_text("Double-click to insert as custom script").double_clicked() {
                            self.cmds.push(Cmd::AddSnippet(i));
                        }
                    }
                });
                ui.add_space(4.0);
                if crate::widgets::secondary_button(ui, &format!("{}  New custom script", icon::CODE)).clicked() {
                    self.cmds.push(Cmd::AddCustomScript);
                }
            });

        // Header + search
        ui.horizontal(|ui| {
            ui.label(egui::RichText::new("Command Library").font(semibold(15.0)));
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                ui.label(
                    it(format!("{}", self.catalog.visible_commands().count()))
                        .size(11.5)
                        .color(p.text_faint),
                );
            });
        });
        ui.add_space(4.0);
        ui.add(
            egui::TextEdit::singleline(&mut self.lib_search)
                .hint_text(it(format!("{}  Filter commands", icon::FUNNEL)))
                .desired_width(f32::INFINITY)
                .margin(egui::Margin::symmetric(10, 6)),
        );
        ui.add_space(6.0);

        let q = self.lib_search.trim().to_lowercase();
        let all: Vec<(String, String, String)> = self
            .catalog
            .visible_commands()
            .filter(|c| q.is_empty() || c.command.to_lowercase().contains(&q) || c.description.to_lowercase().contains(&q))
            .map(|c| (c.command.clone(), c.name.clone(), c.description.clone()))
            .collect();

        egui::ScrollArea::vertical().id_salt("lib").auto_shrink([false, false]).show(ui, |ui| {
            ui.spacing_mut().item_spacing.y = 1.0;
            if q.is_empty() {
                section(ui, "Frequently used");
                for c in COMMON {
                    if let Some((cmd, _, desc)) = all.iter().find(|(x, _, _)| x == c) {
                        self.lib_item(ui, cmd, desc, "common");
                    }
                }
                ui.add_space(6.0);
                section(ui, "All commands");
            } else if all.is_empty() {
                crate::widgets::muted(ui, "No command matches the filter.");
            }
            let mut last_verb = String::new();
            for (cmd, _, desc) in &all {
                let verb = cmd.split('-').next().unwrap_or("").to_string();
                if q.is_empty() && verb != last_verb {
                    ui.add_space(3.0);
                    ui.label(egui::RichText::new(&verb).size(11.0).color(p.text_faint));
                    last_verb = verb;
                }
                self.lib_item(ui, cmd, desc, "all");
            }
        });
    }

    fn lib_item(&mut self, ui: &mut egui::Ui, cmd: &str, desc: &str, salt: &str) {
        let p = self.palette;
        let selected = self.lib_selected.as_deref() == Some(cmd);
        let id = egui::Id::new(("lib", salt, cmd));
        let resp = ui
            .dnd_drag_source(id, DragCommand(cmd.to_string()), |ui| {
                list_row(ui, verb_icon(cmd), cmd, selected, p.accent)
            })
            .inner;
        // dnd_drag_source only senses drags; add click sensing on top
        let resp = ui.interact(resp.rect, id.with("click"), Sense::click());
        if resp.clicked() {
            self.lib_selected = Some(cmd.to_string());
        }
        if resp.double_clicked() {
            self.cmds.push(Cmd::AddCommand(cmd.to_string()));
        }
        let desc = desc.to_string();
        let cmd_s = cmd.to_string();
        resp.on_hover_ui(|ui| {
            ui.set_max_width(340.0);
            ui.label(egui::RichText::new(&cmd_s).monospace().strong());
            ui.add_space(2.0);
            let short: String = desc.chars().take(400).collect();
            ui.label(short);
            ui.add_space(4.0);
            ui.label(egui::RichText::new("Double-click or drag into the sequence").size(11.0).weak());
        })
        .context_menu(|ui| {
            if ui.button(it(format!("{}  Add to sequence", icon::PLUS))).clicked() {
                self.cmds.push(Cmd::AddCommand(cmd.to_string()));
                ui.close();
            }
            if ui.button(it(format!("{}  Online help", icon::QUESTION))).clicked() {
                self.cmds.push(Cmd::OpenUrl(crate::catalog::Catalog::help_url(cmd)));
                ui.close();
            }
        });
    }
}

fn section(ui: &mut egui::Ui, text: &str) {
    let p = crate::widgets::palette(ui);
    ui.add_space(2.0);
    ui.label(egui::RichText::new(text.to_uppercase()).size(11.0).font(semibold(11.0)).color(p.text_muted));
    ui.add_space(2.0);
}

fn list_row(ui: &mut egui::Ui, ic: &str, text: &str, selected: bool, icon_color: egui::Color32) -> egui::Response {
    let p = crate::widgets::palette(ui);
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(ui.available_width(), 26.0), Sense::click());
    if ui.is_rect_visible(rect) {
        if selected {
            ui.painter().rect_filled(rect, CornerRadius::same(6), p.accent_soft);
            let bar = egui::Rect::from_min_size(rect.left_top() + Vec2::new(0.0, 6.0), Vec2::new(3.0, 14.0));
            ui.painter().rect_filled(bar, CornerRadius::same(2), p.accent);
        } else if resp.hovered() {
            ui.painter().rect_filled(rect, CornerRadius::same(6), p.hover);
        }
        ui.painter().text(
            egui::pos2(rect.left() + 16.0, rect.center().y),
            Align2::CENTER_CENTER,
            ic,
            crate::theme::icon_font(15.0),
            if selected { p.accent } else { icon_color.gamma_multiply(0.9) },
        );
        ui.painter().text(
            egui::pos2(rect.left() + 32.0, rect.center().y),
            Align2::LEFT_CENTER,
            text,
            FontId::proportional(13.0),
            if selected { p.accent_text } else { p.text },
        );
    }
    resp
}
