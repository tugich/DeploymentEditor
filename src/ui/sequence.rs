//! The central sequence view: deployment-type switcher and action cards.

use eframe::egui::{self, Align2, CornerRadius, FontId, Sense, Stroke, Vec2};
use egui_phosphor::regular as icon;

use super::{DragAction, DragCommand};
use crate::app::{Cmd, DeploymentEditorApp};
use crate::project::{Action, DeploymentType};
use crate::theme::semibold;
use crate::widgets::{it, chip, code_chip, icon_button, segmented};

impl DeploymentEditorApp {
    pub(super) fn sequence(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;

        // Header -----------------------------------------------------------
        ui.horizontal(|ui| {
            let mut dt = self.dt;
            let items: Vec<(DeploymentType, String)> = DeploymentType::ALL
                .iter()
                .enumerate()
                .map(|(i, d)| (*d, format!("{}  {}", d.as_db(), self.counts[i])))
                .collect();
            if segmented(ui, &mut dt, &items) {
                self.cmds.push(Cmd::SetDeploymentType(dt));
            }
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if icon_button(ui, icon::ARROWS_IN_SIMPLE, "Collapse all").clicked() {
                    self.cmds.push(Cmd::ExpandAll(false));
                }
                if icon_button(ui, icon::ARROWS_OUT_SIMPLE, "Expand all").clicked() {
                    self.cmds.push(Cmd::ExpandAll(true));
                }
                if icon_button(ui, icon::CODE, "Add custom script").clicked() {
                    self.cmds.push(Cmd::AddCustomScript);
                }
            });
        });
        ui.add_space(4.0);
        ui.label(
            egui::RichText::new(match self.dt {
                DeploymentType::Installation => "Steps executed in the Install phase of Invoke-AppDeployToolkit.ps1",
                DeploymentType::Uninstall => "Steps executed in the Uninstall phase of Invoke-AppDeployToolkit.ps1",
                DeploymentType::Repair => "Steps executed in the Repair phase of Invoke-AppDeployToolkit.ps1",
            })
            .size(12.0)
            .color(p.text_muted),
        );
        ui.add_space(4.0);

        // Cards -------------------------------------------------------------
        let dragging_something = egui::DragAndDrop::has_any_payload(ui.ctx());
        let actions = std::mem::take(&mut self.actions);
        let n = actions.len();
        let mut drop_target: Option<(usize, bool)> = None; // (index, above)
        let mut last_rect: Option<egui::Rect> = None;

        let scroll = egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
            ui.spacing_mut().item_spacing.y = 8.0;
            ui.add_space(6.0);
            if actions.is_empty() {
                self.empty_state(ui);
            }
            for (idx, a) in actions.iter().enumerate() {
                let resp = self.action_card(ui, idx, a);
                last_rect = Some(resp.rect);
                if self.scroll_to_selected && self.selected == Some(a.id) {
                    resp.scroll_to_me(Some(egui::Align::Center));
                    self.scroll_to_selected = false;
                }
                if dragging_something {
                    if let Some(pos) = ui.ctx().pointer_latest_pos() {
                        if resp.rect.expand2(Vec2::new(0.0, 4.0)).contains(pos) {
                            let above = pos.y < resp.rect.center().y;
                            drop_target = Some((idx, above));
                            let y = if above { resp.rect.top() - 4.0 } else { resp.rect.bottom() + 4.0 };
                            ui.painter().line_segment(
                                [egui::pos2(resp.rect.left() + 8.0, y), egui::pos2(resp.rect.right() - 8.0, y)],
                                Stroke::new(3.0, p.accent),
                            );
                            ui.painter().circle_filled(egui::pos2(resp.rect.left() + 8.0, y), 4.5, p.accent);
                        }
                    }
                }
            }
            // Space at the bottom so items can be dropped "at the end".
            let (rest, _) = ui.allocate_exact_size(Vec2::new(ui.available_width(), 60.0), Sense::hover());
            if dragging_something && n > 0 {
                if let Some(pos) = ui.ctx().pointer_latest_pos() {
                    if rest.contains(pos) {
                        drop_target = Some((n - 1, false));
                        if let Some(r) = last_rect {
                            let y = r.bottom() + 4.0;
                            ui.painter().line_segment(
                                [egui::pos2(r.left() + 8.0, y), egui::pos2(r.right() - 8.0, y)],
                                Stroke::new(3.0, p.accent),
                            );
                            ui.painter().circle_filled(egui::pos2(r.left() + 8.0, y), 4.5, p.accent);
                        }
                    }
                }
            }
        });

        // Drag ghost: a pill with the name of what is being dragged.
        if dragging_something {
            let label = if let Some(d) = egui::DragAndDrop::payload::<DragAction>(ui.ctx()) {
                actions.iter().find(|a| a.id == d.0).map(|a| format!("{}  {}", icon::DOTS_SIX_VERTICAL, a.name))
            } else {
                egui::DragAndDrop::payload::<DragCommand>(ui.ctx()).map(|c| format!("{}  {}", icon::PLUS_CIRCLE, c.0))
            };
            if let (Some(label), Some(pos)) = (label, ui.ctx().pointer_latest_pos()) {
                let painter = ui.ctx().layer_painter(egui::LayerId::new(egui::Order::Tooltip, egui::Id::new("drag-ghost")));
                let galley = painter.layout_job(crate::widgets::icon_job(&label, crate::theme::semibold(13.0), p.text));
                let rect = egui::Rect::from_min_size(pos + Vec2::new(14.0, 6.0), galley.size() + Vec2::new(24.0, 14.0));
                painter.rect(rect, CornerRadius::same(8), p.panel_bg, Stroke::new(1.5, p.accent), egui::StrokeKind::Inside);
                painter.galley(rect.min + Vec2::new(12.0, 7.0), galley, p.text);
                ui.ctx().set_cursor_icon(egui::CursorIcon::Grabbing);
            }
        }

        // Handle drops --------------------------------------------------------
        let released = ui.input(|i| i.pointer.any_released());
        if released && scroll.inner_rect.contains(ui.ctx().pointer_interact_pos().unwrap_or_default()) {
            // NB: take_payload() clears the payload even if the type does not match,
            // so check the type first.
            let is_action = egui::DragAndDrop::payload::<DragAction>(ui.ctx()).is_some();
            let is_command = egui::DragAndDrop::payload::<DragCommand>(ui.ctx()).is_some();
            if let Some(payload) = is_action.then(|| egui::DragAndDrop::take_payload::<DragAction>(ui.ctx())).flatten() {
                if let Some((idx, above)) = drop_target {
                    let from = actions.iter().position(|a| a.id == payload.0).unwrap_or(idx);
                    let mut target = if above { idx } else { idx + 1 };
                    if from < target {
                        target -= 1;
                    }
                    self.cmds.push(Cmd::MoveTo(payload.0, target.min(n.saturating_sub(1))));
                }
            } else if let Some(payload) = is_command.then(|| egui::DragAndDrop::take_payload::<DragCommand>(ui.ctx())).flatten() {
                self.cmds.push(Cmd::AddCommand(payload.0.clone()));
                if let Some((idx, above)) = drop_target {
                    let target = if above { idx } else { idx + 1 };
                    if target < n {
                        // will be appended at n; move it into place after it exists
                        self.pending_insert_at = Some(target);
                    }
                }
            }
        }
        self.actions = actions;
    }

    fn empty_state(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        ui.add_space(40.0);
        ui.vertical_centered(|ui| {
            ui.label(it(icon::STACK_PLUS).size(46.0).color(p.text_faint));
            ui.add_space(6.0);
            ui.label(egui::RichText::new("No actions yet").font(semibold(17.0)));
            ui.label(
                egui::RichText::new(
                    "Double-click a command in the library, drag it here,\nuse the Insert tab — or drop an EXE / MSI / PS1 file onto the window.",
                )
                .color(p.text_muted),
            );
            ui.add_space(10.0);
            ui.horizontal(|ui| {
                let w = 330.0;
                ui.add_space((ui.available_width() - w).max(0.0) / 2.0);
                if crate::widgets::primary_button(ui, &format!("{}  Run an installer", icon::APP_WINDOW)).clicked() {
                    self.cmds.push(Cmd::AddCommandWith(
                        "Start-ADTProcess".into(),
                        "Run installer".into(),
                        vec![("FilePath".into(), "setup.exe".into()), ("ArgumentList".into(), "/S".into())],
                    ));
                }
                if crate::widgets::secondary_button(ui, &format!("{}  Custom script", icon::CODE)).clicked() {
                    self.cmds.push(Cmd::AddCustomScript);
                }
            });
        });
    }

    fn action_card(&mut self, ui: &mut egui::Ui, idx: usize, a: &Action) -> egui::Response {
        let p = self.palette;
        let selected = self.selected == Some(a.id);
        let expanded = self.expanded.contains(&a.id);
        // While editing, show the draft so the card updates live.
        let draft_copy = if selected { self.draft.as_ref().filter(|d| d.id == a.id).cloned() } else { None };
        let a = draft_copy.as_ref().unwrap_or(a);

        let scope = ui.scope_builder(egui::UiBuilder::new().sense(Sense::click()).id_salt(("card", a.id)), |ui| {
            let hovered = ui.response().hovered();
            let fill = if selected {
                if p.dark { p.accent_soft } else { p.card_bg }
            } else if hovered {
                p.card_hover
            } else {
                p.card_bg
            };
            let stroke = if selected { Stroke::new(1.5, p.accent) } else { Stroke::new(1.0, p.border) };
            let frame = egui::Frame::new()
                .fill(fill)
                .stroke(stroke)
                .corner_radius(CornerRadius::same(10))
                .inner_margin(egui::Margin { left: 6, right: 12, top: 10, bottom: 10 });
            frame.show(ui, |ui| {
                ui.set_width(ui.available_width());
                if a.disabled {
                    ui.set_opacity(0.55);
                }
                ui.horizontal_top(|ui| {
                    ui.spacing_mut().item_spacing.x = 8.0;
                    // drag handle
                    let handle_id = egui::Id::new(("drag", a.id));
                    ui.dnd_drag_source(handle_id, DragAction(a.id), |ui| {
                        let (r, resp) = ui.allocate_exact_size(Vec2::new(16.0, 36.0), Sense::hover());
                        let col = if resp.hovered() { p.text_muted } else { p.text_faint };
                        ui.painter().text(r.center(), Align2::CENTER_CENTER, icon::DOTS_SIX_VERTICAL, crate::theme::icon_font(16.0), col);
                    })
                    .response
                    .on_hover_cursor(egui::CursorIcon::Grab)
                    .on_hover_text("Drag to reorder");

                    // step badge with status icon
                    let (badge, _) = ui.allocate_exact_size(Vec2::new(36.0, 36.0), Sense::hover());
                    let (ic, col, bg) = if a.disabled {
                        (icon::PROHIBIT, p.text_faint, p.hover)
                    } else if a.is_custom_script() {
                        (icon::CODE, p.script, p.script.gamma_multiply(0.14))
                    } else if !a.conditions.trim().is_empty() {
                        (icon::GIT_BRANCH, p.info, p.info.gamma_multiply(0.14))
                    } else {
                        (icon::CHECK_CIRCLE, p.success, p.success.gamma_multiply(0.13))
                    };
                    ui.painter().rect_filled(badge, CornerRadius::same(8), bg);
                    ui.painter().text(badge.center(), Align2::CENTER_CENTER, ic, crate::theme::icon_font(19.0), col);

                    ui.vertical(|ui| {
                        ui.spacing_mut().item_spacing.y = 4.0;
                        ui.horizontal_wrapped(|ui| {
                            ui.spacing_mut().item_spacing.x = 6.0;
                            ui.label(it(format!("{}.", idx + 1)).font(semibold(14.0)).color(p.text_faint));
                            let name = if a.name.trim().is_empty() { "(unnamed action)" } else { a.name.as_str() };
                            ui.label(egui::RichText::new(name).font(semibold(14.5)).color(p.text));
                            if !a.disabled {
                                let problems = self.issues_for(a.id);
                                let errors = problems.iter().filter(|i| i.severity == crate::validate::Severity::Error).count();
                                if !problems.is_empty() {
                                    let (fg, label) = if errors > 0 {
                                        (p.danger, format!("{}  {} error{}", icon::WARNING_CIRCLE, errors, if errors == 1 { "" } else { "s" }))
                                    } else {
                                        (p.warning, format!("{}  {} warning{}", icon::WARNING, problems.len(), if problems.len() == 1 { "" } else { "s" }))
                                    };
                                    chip(ui, &label, fg, fg.gamma_multiply(0.13)).on_hover_ui(|ui| {
                                        ui.set_max_width(360.0);
                                        for i in &problems {
                                            ui.label(format!("• {}", i.message));
                                        }
                                    });
                                }
                            }
                            if a.disabled {
                                chip(ui, "Disabled", p.text_muted, p.hover);
                            }
                            if a.continue_on_error {
                                chip(ui, "Continue on error", p.warning, p.warning.gamma_multiply(0.14));
                            }
                            if !a.variable_name.trim().is_empty() {
                                chip(ui, &format!("{} {}", icon::ARROW_RIGHT, a.variable_name), p.info, p.info.gamma_multiply(0.12));
                            }
                            if !a.conditions.trim().is_empty() {
                                chip(ui, &format!("if ({})", truncate(&a.conditions, 40)), p.info, p.info.gamma_multiply(0.12))
                                    .on_hover_text(format!("Runs only if: {}\n\n{}", crate::condition::describe(&a.conditions), a.conditions));
                            }
                        });
                        ui.horizontal(|ui| {
                            let cmd_label = if a.is_custom_script() { "PowerShell" } else { a.command.as_str() };
                            code_chip(ui, cmd_label);
                            if !a.description.trim().is_empty() {
                                ui.add(
                                    egui::Label::new(
                                        egui::RichText::new(a.description.lines().next().unwrap_or("")).size(12.5).color(p.text_muted),
                                    )
                                    .truncate(),
                                );
                            }
                        });
                    });

                    ui.with_layout(egui::Layout::right_to_left(egui::Align::TOP), |ui| {
                        let chevron = if expanded { icon::CARET_UP } else { icon::CARET_DOWN };
                        if icon_button(ui, chevron, if expanded { "Hide parameters" } else { "Show parameters" }).clicked() {
                            if expanded {
                                self.expanded.remove(&a.id);
                            } else {
                                self.expanded.insert(a.id);
                            }
                        }
                        let summary = param_summary(a, &self.catalog);
                        if !summary.is_empty() && !expanded {
                            ui.label(egui::RichText::new(summary).size(11.5).color(p.text_faint));
                        }
                    });
                });

                if expanded {
                    ui.add_space(6.0);
                    ui.indent(("params", a.id), |ui| {
                        ui.add_space(0.0);
                        self.card_details(ui, a);
                    });
                }
            });
        });

        let resp = scope.response;
        if resp.clicked() {
            self.cmds.push(Cmd::Select(Some(a.id)));
        }
        if resp.double_clicked() {
            if expanded {
                self.expanded.remove(&a.id);
            } else {
                self.expanded.insert(a.id);
            }
        }
        let id = a.id;
        let disabled = a.disabled;
        resp.context_menu(|ui| {
            ui.set_min_width(200.0);
            if !selected {
                self.cmds.push(Cmd::Select(Some(id)));
            }
            if ui.button(it(format!("{}  Move up", icon::ARROW_UP))).clicked() {
                self.cmds.push(Cmd::Move(-1));
                ui.close();
            }
            if ui.button(it(format!("{}  Move down", icon::ARROW_DOWN))).clicked() {
                self.cmds.push(Cmd::Move(1));
                ui.close();
            }
            ui.separator();
            if ui.button(it(format!("{}  Duplicate", icon::COPY))).clicked() {
                self.cmds.push(Cmd::Duplicate);
                ui.close();
            }
            let label = if disabled { format!("{}  Enable", icon::EYE) } else { format!("{}  Disable", icon::EYE_SLASH) };
            if ui.button(it(label)).clicked() {
                self.cmds.push(Cmd::ToggleDisable(Some(id)));
                ui.close();
            }
            if ui.button(it(format!("{}  Command help", icon::QUESTION))).clicked() {
                self.cmds.push(Cmd::CommandHelp);
                ui.close();
            }
            ui.separator();
            if ui.button(it(format!("{}  Remove…", icon::TRASH)).color(self.palette.danger)).clicked() {
                self.cmds.push(Cmd::RemoveSelected);
                ui.close();
            }
        });
        resp
    }

    fn card_details(&self, ui: &mut egui::Ui, a: &Action) {
        let p = self.palette;
        if a.is_custom_script() {
            let script = a.value("Script").replace("\r\n", "\n");
            let preview: String = script.lines().take(12).collect::<Vec<_>>().join("\n");
            let more = script.lines().count().saturating_sub(12);
            egui::Frame::new()
                .fill(ui.visuals().code_bg_color)
                .corner_radius(CornerRadius::same(6))
                .inner_margin(egui::Margin::same(10))
                .show(ui, |ui| {
                    ui.set_width(ui.available_width());
                    if preview.trim().is_empty() {
                        ui.label(egui::RichText::new("# empty script — select the action to edit it").monospace().color(p.text_faint));
                    } else {
                        let job = crate::highlight::highlight_powershell(&preview, FontId::monospace(12.5), p.dark);
                        ui.label(job);
                    }
                    if more > 0 {
                        ui.label(it(format!("… {more} more lines")).size(11.5).color(p.text_faint));
                    }
                });
            return;
        }
        let set: Vec<(String, String, bool)> = a
            .values
            .iter()
            .filter(|(_, v)| !v.trim().is_empty())
            .filter_map(|(k, v)| {
                let info = self.catalog.parameter(&a.command, k);
                let is_switch = info.map(|i| i.ptype == "SwitchParameter").unwrap_or(false);
                if is_switch && v != "1" {
                    None
                } else {
                    Some((k.clone(), v.clone(), is_switch))
                }
            })
            .collect();
        if set.is_empty() {
            ui.label(egui::RichText::new("No parameters set — select the action and open Parameters.").size(12.5).color(p.text_faint));
            return;
        }
        egui::Grid::new(("grid", a.id)).num_columns(2).spacing(Vec2::new(16.0, 4.0)).show(ui, |ui| {
            for (k, v, is_switch) in set {
                ui.label(it(format!("-{k}")).monospace().size(12.5).color(p.text_muted));
                if is_switch {
                    ui.label(it(format!("{}  on", icon::CHECK)).size(12.5).color(p.success));
                } else {
                    ui.add(egui::Label::new(egui::RichText::new(v).monospace().size(12.5)).wrap());
                }
                ui.end_row();
            }
        });
    }
}

fn truncate(s: &str, n: usize) -> String {
    if s.chars().count() <= n {
        s.to_string()
    } else {
        let mut t: String = s.chars().take(n).collect();
        t.push('…');
        t
    }
}

fn param_summary(a: &Action, catalog: &crate::catalog::Catalog) -> String {
    if a.is_custom_script() {
        let lines = a.value("Script").lines().filter(|l| !l.trim().is_empty()).count();
        return format!("{lines} lines");
    }
    let n = a
        .values
        .iter()
        .filter(|(k, v)| {
            !v.trim().is_empty()
                && !(catalog.parameter(&a.command, k).map(|i| i.ptype == "SwitchParameter").unwrap_or(false) && v != "1")
        })
        .count();
    match n {
        0 => String::new(),
        1 => "1 parameter".into(),
        n => format!("{n} parameters"),
    }
}
