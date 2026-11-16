//! Right pane: properties of the selected action (options + parameters).

use eframe::egui::{self, CornerRadius, FontId, Stroke, Vec2};
use egui_phosphor::regular as icon;

use crate::app::{Cmd, DeploymentEditorApp, InspectorTab};
use crate::theme::semibold;
use crate::widgets::{it, field_label, icon_button, muted, primary_button, secondary_button, segmented};

impl DeploymentEditorApp {
    pub(super) fn inspector(&mut self, ui: &mut egui::Ui) {
        let p = self.palette;
        let Some(draft) = self.draft.clone() else {
            ui.add_space(60.0);
            ui.vertical_centered(|ui| {
                ui.label(it(icon::CURSOR_CLICK).size(40.0).color(p.text_faint));
                ui.add_space(6.0);
                ui.label(egui::RichText::new("No action selected").font(semibold(15.0)));
                muted(ui, "Select a step in the sequence to edit\nits options and parameters.");
            });
            return;
        };

        // Header -----------------------------------------------------------
        ui.horizontal(|ui| {
            let (ic, col) = if draft.is_custom_script() { (icon::CODE, p.script) } else { (crate::ui::library::verb_icon(&draft.command), p.accent) };
            let (r, _) = ui.allocate_exact_size(Vec2::splat(38.0), egui::Sense::hover());
            ui.painter().rect_filled(r, CornerRadius::same(9), col.gamma_multiply(0.14));
            ui.painter().text(r.center(), egui::Align2::CENTER_CENTER, ic, crate::theme::icon_font(21.0), col);
            ui.vertical(|ui| {
                ui.spacing_mut().item_spacing.y = 0.0;
                ui.add(egui::Label::new(egui::RichText::new(&draft.name).font(semibold(15.5))).truncate());
                let cmd = if draft.is_custom_script() { "Custom PowerShell script".to_string() } else { draft.command.clone() };
                ui.label(egui::RichText::new(cmd).monospace().size(12.0).color(p.text_muted));
            });
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if !draft.is_custom_script() && icon_button(ui, icon::QUESTION, "Online help for this command (F1)").clicked() {
                    self.cmds.push(Cmd::CommandHelp);
                }
            });
        });
        ui.add_space(10.0);
        let mut tab = self.inspector_tab;
        let params_label = if draft.is_custom_script() { "Script".to_string() } else { "Parameters".to_string() };
        if segmented(ui, &mut tab, &[(InspectorTab::Options, "Options".into()), (InspectorTab::Parameters, params_label)]) {
            self.inspector_tab = tab;
        }
        ui.add_space(8.0);

        // Footer (save bar) ---------------------------------------------------
        egui::Panel::bottom("inspector-footer")
            .show_separator_line(true)
            .frame(egui::Frame::new().inner_margin(egui::Margin { left: 0, right: 0, top: 10, bottom: 2 }))
            .show(ui, |ui| {
                ui.horizontal(|ui| {
                    ui.add_enabled_ui(self.dirty, |ui| {
                        if primary_button(ui, &format!("{}  Save", icon::FLOPPY_DISK)).on_hover_text("Ctrl+S").clicked() {
                            self.cmds.push(Cmd::Save);
                        }
                        if secondary_button(ui, "Revert").clicked() {
                            self.dirty = false;
                            self.draft = self.actions.iter().find(|a| Some(a.id) == self.selected).cloned();
                        }
                    });
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if self.dirty {
                            ui.label(it(format!("{}  Unsaved", icon::CIRCLE)).size(12.0).color(p.warning));
                        } else {
                            ui.label(it(format!("{}  Saved", icon::CHECK)).size(12.0).color(p.text_faint));
                        }
                    });
                });
                ui.label(egui::RichText::new("Changes are also saved when you select another step.").size(11.0).color(p.text_faint));
            });

        // Body ---------------------------------------------------------------
        let mut d = draft;
        let mut changed = false;
        egui::ScrollArea::vertical().id_salt("insp").auto_shrink([false, false]).show(ui, |ui| {
            ui.spacing_mut().item_spacing.y = 6.0;
            match self.inspector_tab {
                InspectorTab::Options => changed |= self.options_tab(ui, &mut d),
                InspectorTab::Parameters => {
                    if d.is_custom_script() {
                        changed |= self.script_tab(ui, &mut d);
                    } else {
                        changed |= self.params_tab(ui, &mut d);
                    }
                }
            }
        });
        if changed {
            self.draft = Some(d);
            self.dirty = true;
        }
    }

    fn options_tab(&mut self, ui: &mut egui::Ui, d: &mut crate::project::Action) -> bool {
        let p = self.palette;
        let mut changed = false;
        field_label(ui, "Name", None);
        changed |= ui
            .add(egui::TextEdit::singleline(&mut d.name).desired_width(f32::INFINITY).margin(egui::Margin::symmetric(8, 6)))
            .changed();
        ui.add_space(4.0);
        field_label(ui, "Description", Some("written as a # comment"));
        changed |= ui
            .add(egui::TextEdit::multiline(&mut d.description).desired_rows(3).desired_width(f32::INFINITY))
            .changed();
        ui.add_space(4.0);
        field_label(ui, "Store result in variable", Some("e.g. $result"));
        changed |= ui
            .add(
                egui::TextEdit::singleline(&mut d.variable_name)
                    .hint_text("$variable")
                    .font(egui::TextStyle::Monospace)
                    .desired_width(f32::INFINITY)
                    .margin(egui::Margin::symmetric(8, 6)),
            )
            .changed();
        ui.add_space(4.0);
        field_label(ui, "Run only if (condition)", Some("PowerShell expression or the builder below"));
        egui::ScrollArea::vertical().id_salt("cond-text").max_height(64.0).show(ui, |ui| {
            changed |= ui
                .add(
                    egui::TextEdit::multiline(&mut d.conditions)
                        .hint_text("e.g. $Is64Bit -and $envOSVersionBuild -ge 22631")
                        .font(egui::TextStyle::Monospace)
                        .desired_rows(2)
                        .desired_width(f32::INFINITY),
                )
                .changed();
        });
        ui.add_space(2.0);
        changed |= self.condition_builder(ui, d);
        ui.add_space(8.0);
        ui.separator();
        ui.add_space(4.0);
        changed |= ui
            .checkbox(&mut d.continue_on_error, "Continue on error")
            .on_hover_text("Adds -ErrorAction SilentlyContinue")
            .changed();
        changed |= ui.checkbox(&mut d.disabled, "Disable this action").on_hover_text("Disabled steps are not generated").changed();
        ui.add_space(10.0);
        egui::Frame::new()
            .fill(p.hover)
            .corner_radius(CornerRadius::same(8))
            .inner_margin(egui::Margin::same(10))
            .show(ui, |ui| {
                ui.set_width(ui.available_width());
                ui.label(it(format!("{}  Generated line", icon::CODE)).font(semibold(12.5)));
                let mut preview = d.clone();
                preview.disabled = false;
                let line = crate::generator::build_script(std::slice::from_ref(&preview), &self.catalog);
                if d.disabled {
                    ui.label(it("Disabled — this step is skipped when the script is generated.").size(11.5).color(p.warning));
                }
                let line: String = line
                    .trim_start_matches("\r\n")
                    .split("\r\n")
                    .map(|l| l.strip_prefix("    ").unwrap_or(l))
                    .collect::<Vec<_>>()
                    .join("\n");
                let job = crate::highlight::highlight_powershell(&line, FontId::monospace(12.0), p.dark);
                ui.add(egui::Label::new(job).wrap());
            });
        changed
    }

    fn params_tab(&mut self, ui: &mut egui::Ui, d: &mut crate::project::Action) -> bool {
        let p = self.palette;
        let mut changed = false;
        let params = self.catalog.parameters(&d.command).to_vec();
        if params.is_empty() {
            muted(ui, "This command has no parameters in the PSADT database.");
            return false;
        }
        ui.horizontal(|ui| {
            ui.add(
                egui::TextEdit::singleline(&mut self.param_search)
                    .hint_text(it(format!("{}  Filter parameters", icon::FUNNEL)))
                    .desired_width(ui.available_width() - 110.0)
                    .margin(egui::Margin::symmetric(8, 5)),
            );
            ui.checkbox(&mut self.only_set_params, "Only set");
        });
        ui.add_space(6.0);
        let q = self.param_search.trim().to_lowercase();
        let mut shown = 0;
        for param in &params {
            let current = d.value(&param.name).to_string();
            let is_set = !current.trim().is_empty() && !(param.is_checkbox() && current != "1");
            if self.only_set_params && !is_set {
                continue;
            }
            if !q.is_empty() && !param.name.to_lowercase().contains(&q) && !param.description.to_lowercase().contains(&q) {
                continue;
            }
            shown += 1;
            egui::Frame::new()
                .stroke(Stroke::new(1.0, if is_set { p.accent.gamma_multiply(0.5) } else { p.border }))
                .fill(if is_set { p.accent_soft.gamma_multiply(if p.dark { 0.6 } else { 0.5 }) } else { p.panel_bg })
                .corner_radius(CornerRadius::same(8))
                .inner_margin(egui::Margin::symmetric(10, 8))
                .show(ui, |ui| {
                    ui.set_width(ui.available_width());
                    if param.is_checkbox() {
                        let mut on = current == "1" || current.eq_ignore_ascii_case("Enabled");
                        ui.horizontal(|ui| {
                            if ui.checkbox(&mut on, egui::RichText::new(&param.name).font(semibold(13.0))).changed() {
                                d.set_value(&param.name, if on { "1" } else { "0" });
                                changed = true;
                            }
                            if param.required {
                                ui.label(egui::RichText::new("*").color(p.danger)).on_hover_text("Mandatory in at least one parameter set");
                            }
                            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                ui.label(egui::RichText::new("switch").size(11.0).color(p.text_faint));
                            });
                        });
                    } else {
                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new(&param.name).font(semibold(13.0)));
                            if param.required {
                                ui.label(egui::RichText::new("*").color(p.danger)).on_hover_text("Mandatory in at least one parameter set");
                            }
                            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                ui.label(egui::RichText::new(&param.ptype).size(11.0).color(p.text_faint));
                            });
                        });
                        let mut v = current.clone();
                        let hint = type_hint(&param.ptype);
                        if ui
                            .add(
                                egui::TextEdit::singleline(&mut v)
                                    .hint_text(hint)
                                    .font(egui::TextStyle::Monospace)
                                    .desired_width(f32::INFINITY)
                                    .margin(egui::Margin::symmetric(8, 5)),
                            )
                            .changed()
                        {
                            d.set_value(&param.name, v);
                            changed = true;
                        }
                    }
                    if !param.description.is_empty() {
                        let short: String = param.description.chars().take(160).collect();
                        let more = param.description.chars().count() > 160;
                        let r = ui.label(
                            egui::RichText::new(if more { format!("{short}…") } else { short })
                                .size(11.5)
                                .color(p.text_muted),
                        );
                        if more {
                            r.on_hover_text(&param.description);
                        }
                    }
                });
            ui.add_space(2.0);
        }
        if shown == 0 {
            muted(ui, "No parameters match the filter.");
        }
        changed
    }

    fn script_tab(&mut self, ui: &mut egui::Ui, d: &mut crate::project::Action) -> bool {
        let p = self.palette;
        let mut script = d.value("Script").replace("\r\n", "\n");
        ui.horizontal(|ui| {
            field_label(ui, "PowerShell", None);
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if ui.button(it(format!("{}  Large editor", icon::ARROWS_OUT))).clicked() {
                    self.script_editor_open = true;
                }
            });
        });
        let dark = p.dark;
        let mut layouter = move |ui: &egui::Ui, text: &dyn egui::TextBuffer, wrap: f32| {
            let mut job = crate::highlight::highlight_powershell(text.as_str(), FontId::monospace(12.5), dark);
            job.wrap.max_width = wrap;
            ui.fonts_mut(|f| f.layout_job(job))
        };
        let r = ui.add(
            egui::TextEdit::multiline(&mut script)
                .code_editor()
                .desired_rows(22)
                .desired_width(f32::INFINITY)
                .layouter(&mut layouter),
        );
        muted(ui, "Indentation is added automatically when the script is generated.");
        if r.changed() {
            d.set_value("Script", script);
            return true;
        }
        false
    }

    /// A resizable window with a big code editor for custom scripts.
    pub(super) fn script_editor_window(&mut self, ctx: &egui::Context) {
        if !self.script_editor_open {
            return;
        }
        let Some(mut d) = self.draft.clone() else {
            self.script_editor_open = false;
            return;
        };
        if !d.is_custom_script() {
            self.script_editor_open = false;
            return;
        }
        let p = self.palette;
        let mut open = true;
        let mut changed = false;
        let mut save = false;
        egui::Window::new(it(format!("{}  {}", icon::CODE, d.name)))
            .open(&mut open)
            .pivot(egui::Align2::CENTER_CENTER)
            .default_pos(ctx.content_rect().center())
            .default_size(Vec2::new(860.0, 560.0))
            .resizable(true)
            .collapsible(false)
            .show(ctx, |ui| {
                egui::Panel::bottom("se-footer").show_separator_line(false).show(ui, |ui| {
                    ui.add_space(6.0);
                    ui.horizontal(|ui| {
                        if primary_button(ui, &format!("{}  Save", icon::FLOPPY_DISK)).clicked() {
                            save = true;
                        }
                        muted(ui, "Ctrl+S saves too. PSADT variables like $adtSession.DirFiles are available.");
                    });
                });
                let mut script = d.value("Script").replace("\r\n", "\n");
                let dark = p.dark;
                let mut layouter = move |ui: &egui::Ui, text: &dyn egui::TextBuffer, wrap: f32| {
                    let mut job = crate::highlight::highlight_powershell(text.as_str(), FontId::monospace(13.5), dark);
                    job.wrap.max_width = wrap;
                    ui.fonts_mut(|f| f.layout_job(job))
                };
                egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
                    if ui
                        .add_sized(
                            ui.available_size(),
                            egui::TextEdit::multiline(&mut script)
                                .code_editor()
                                .desired_width(f32::INFINITY)
                                .layouter(&mut layouter),
                        )
                        .changed()
                    {
                        d.set_value("Script", script);
                        changed = true;
                    }
                });
            });
        if changed {
            self.draft = Some(d);
            self.dirty = true;
        }
        if save {
            self.cmds.push(Cmd::Save);
        }
        if !open {
            self.script_editor_open = false;
        }
    }
}

fn type_hint(t: &str) -> String {
    match t {
        "String[]" => "text".into(),
        "Int32" | "UInt32" | "Int64" => "number".into(),
        "Int32[]" => "e.g. 0, 3010".into(),
        "TimeSpan" => "e.g. (New-TimeSpan -Minutes 5)".into(),
        "ScriptBlock" => "{ … }".into(),
        "Hashtable" => "@{ Key = 'Value' }".into(),
        "Guid" => "{00000000-0000-0000-0000-000000000000}".into(),
        "" | "String" | "Object" => "value".into(),
        other => other.to_string(),
    }
}
