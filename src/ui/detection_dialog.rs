//! Dialog for the Intune detection rule of the project.

use eframe::egui::{self, CornerRadius, FontId, Stroke};
use egui_phosphor::regular as icon;

use crate::app::{DeploymentEditorApp, DetectionForm};
use crate::detection::{DetectionRule, FileCheck, Operator, RegistryCheck, RuleKind};
use crate::theme::semibold;
use crate::widgets::{field_label, it, muted, primary_button, secondary_button, segmented};

pub enum Outcome {
    None,
    Save,
    Remove,
    Close,
}

fn text_field(ui: &mut egui::Ui, value: &mut String, hint: &str, mono: bool) -> bool {
    let mut te = egui::TextEdit::singleline(value)
        .hint_text(hint)
        .desired_width(f32::INFINITY)
        .margin(egui::Margin::symmetric(8, 6));
    if mono {
        te = te.font(egui::TextStyle::Monospace);
    }
    ui.add(te).changed()
}

fn operator_combo(ui: &mut egui::Ui, id: &str, op: &mut Operator) {
    egui::ComboBox::from_id_salt(id)
        .selected_text(format!("{}  {}", op.symbol(), op.graph()))
        .width(200.0)
        .show_ui(ui, |ui| {
            for o in Operator::ALL {
                ui.selectable_value(op, o, format!("{}  {}", o.symbol(), o.graph()));
            }
        });
}

impl DeploymentEditorApp {
    pub(super) fn detection_form(&mut self, ui: &mut egui::Ui, f: &mut DetectionForm) -> Outcome {
        let p = self.palette;
        let mut outcome = Outcome::None;
        ui.set_width(640.0);
        ui.label(it(format!("{}  Detection rule", icon::MAGNIFYING_GLASS_PLUS)).font(semibold(18.0)));
        muted(ui, "How Intune checks whether the app is installed. Used by Upload to Intune and in the PDF documentation.");
        ui.add_space(10.0);

        egui::ScrollArea::vertical().max_height(560.0).show(ui, |ui| {
            // Suggestions -------------------------------------------------
            if !f.suggestions.is_empty() {
                field_label(ui, "Suggestions from this project", None);
                ui.add_space(2.0);
                for (i, s) in f.suggestions.iter().enumerate() {
                    let active = f.rule.as_ref() == Some(&s.rule);
                    egui::Frame::new()
                        .fill(if active { p.accent_soft } else { p.card_bg })
                        .stroke(Stroke::new(1.0, if active { p.accent } else { p.border }))
                        .corner_radius(CornerRadius::same(8))
                        .inner_margin(egui::Margin::symmetric(10, 7))
                        .show(ui, |ui| {
                            ui.set_width(ui.available_width());
                            ui.horizontal(|ui| {
                                let w = (ui.available_width() - 80.0).max(100.0);
                                ui.vertical(|ui| {
                                    ui.set_width(w);
                                    ui.label(egui::RichText::new(&s.title).font(semibold(13.0)));
                                    ui.add(egui::Label::new(egui::RichText::new(&s.reason).size(12.0).color(p.text_muted)).wrap());
                                });
                                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                    if active {
                                        ui.label(it(format!("{}  In use", icon::CHECK)).size(12.0).color(p.accent_text));
                                    } else if secondary_button(ui, "Use").clicked() {
                                        f.rule = Some(f.suggestions[i].rule.clone());
                                    }
                                });
                            });
                        });
                    ui.add_space(3.0);
                }
                ui.add_space(8.0);
            }

            // Editor ----------------------------------------------------------
            field_label(ui, "Rule type", None);
            // `None` = nothing selected yet, so every segment is clickable.
            let mut kind: Option<RuleKind> = f.rule.as_ref().map(|r| r.kind());
            let items: Vec<(Option<RuleKind>, String)> = RuleKind::ALL.iter().map(|k| (Some(*k), k.label().to_string())).collect();
            if segmented(ui, &mut kind, &items) {
                f.rule = kind.map(|k| k.empty());
            }
            if f.rule.is_none() {
                ui.add_space(6.0);
                egui::Frame::new().fill(p.hover).corner_radius(CornerRadius::same(8)).inner_margin(egui::Margin::same(12)).show(ui, |ui| {
                    ui.set_width(ui.available_width());
                    ui.label("No detection rule is configured yet. Pick a suggestion or a rule type.");
                    if secondary_button(ui, &format!("{}  Start with an MSI rule", icon::PLUS)).clicked() {
                        f.rule = Some(RuleKind::Msi.empty());
                    }
                });
            }
            ui.add_space(8.0);

            if let Some(rule) = f.rule.as_mut() {
                match rule {
                    DetectionRule::Msi { product_code, version_operator, product_version } => {
                        field_label(ui, "Product code", Some("{GUID} of the installed MSI"));
                        text_field(ui, product_code, "{00000000-0000-0000-0000-000000000000}", true);
                        ui.add_space(4.0);
                        let mut check_version = *version_operator != Operator::NotConfigured;
                        if ui.checkbox(&mut check_version, "Also check the product version").changed() {
                            *version_operator = if check_version { Operator::GreaterThanOrEqual } else { Operator::NotConfigured };
                        }
                        if check_version {
                            ui.horizontal(|ui| {
                                operator_combo(ui, "msi-op", version_operator);
                                ui.add(egui::TextEdit::singleline(product_version).hint_text("1.0.0").font(egui::TextStyle::Monospace).desired_width(200.0));
                            });
                        }
                    }
                    DetectionRule::File { path, file_or_folder, check, operator, value, check_32bit } => {
                        field_label(ui, "Folder path", Some("environment variables like %ProgramFiles% are allowed"));
                        text_field(ui, path, "%ProgramFiles%\\Vendor\\App", true);
                        field_label(ui, "File or folder name", None);
                        text_field(ui, file_or_folder, "app.exe", true);
                        field_label(ui, "Detection method", None);
                        egui::ComboBox::from_id_salt("file-check").selected_text(check.label()).width(260.0).show_ui(ui, |ui| {
                            for c in FileCheck::ALL {
                                ui.selectable_value(check, c, c.label());
                            }
                        });
                        if check.needs_value() {
                            if *operator == Operator::NotConfigured {
                                *operator = Operator::GreaterThanOrEqual;
                            }
                            ui.horizontal(|ui| {
                                operator_combo(ui, "file-op", operator);
                                let hint = if *check == FileCheck::Version { "1.0.0.0" } else { "10" };
                                ui.add(egui::TextEdit::singleline(value).hint_text(hint).font(egui::TextStyle::Monospace).desired_width(200.0));
                            });
                        }
                        ui.checkbox(check_32bit, "Associated with a 32-bit app on 64-bit clients");
                    }
                    DetectionRule::Registry { key_path, value_name, check, operator, value, check_32bit } => {
                        field_label(ui, "Key path", None);
                        text_field(ui, key_path, "HKEY_LOCAL_MACHINE\\SOFTWARE\\Vendor\\App", true);
                        field_label(ui, "Value name", Some("empty = check the key itself"));
                        text_field(ui, value_name, "DisplayVersion", true);
                        field_label(ui, "Detection method", None);
                        egui::ComboBox::from_id_salt("reg-check").selected_text(check.label()).width(260.0).show_ui(ui, |ui| {
                            for c in RegistryCheck::ALL {
                                ui.selectable_value(check, c, c.label());
                            }
                        });
                        if check.needs_value() {
                            if *operator == Operator::NotConfigured {
                                *operator = Operator::Equal;
                            }
                            ui.horizontal(|ui| {
                                operator_combo(ui, "reg-op", operator);
                                ui.add(egui::TextEdit::singleline(value).font(egui::TextStyle::Monospace).desired_width(200.0));
                            });
                        }
                        ui.checkbox(check_32bit, "Associated with a 32-bit app on 64-bit clients");
                    }
                    DetectionRule::Script { script, run_as_32bit, enforce_signature_check } => {
                        field_label(ui, "PowerShell script", Some("detected = exit code 0 and output on STDOUT"));
                        let dark = p.dark;
                        let mut layouter = move |ui: &egui::Ui, text: &dyn egui::TextBuffer, wrap: f32| {
                            let mut job = crate::highlight::highlight_powershell(text.as_str(), FontId::monospace(12.5), dark);
                            job.wrap.max_width = wrap;
                            ui.fonts_mut(|f| f.layout_job(job))
                        };
                        let mut text = script.replace("\r\n", "\n");
                        if ui
                            .add(
                                egui::TextEdit::multiline(&mut text)
                                    .code_editor()
                                    .desired_rows(14)
                                    .desired_width(f32::INFINITY)
                                    .layouter(&mut layouter),
                            )
                            .changed()
                        {
                            *script = text;
                        }
                        ui.checkbox(run_as_32bit, "Run script as 32-bit process on 64-bit clients");
                        ui.checkbox(enforce_signature_check, "Enforce script signature check");
                    }
                }

                let problems = rule.problems();
                ui.add_space(6.0);
                if problems.is_empty() {
                    ui.label(it(format!("{}  {}", icon::CHECK_CIRCLE, rule.summary())).size(12.5).color(p.success));
                } else {
                    for pr in &problems {
                        ui.label(it(format!("{}  {pr}", icon::WARNING)).size(12.5).color(p.warning));
                    }
                }
            }
        });

        ui.add_space(12.0);
        ui.horizontal(|ui| {
            let valid = f.rule.as_ref().map(|r| r.problems().is_empty()).unwrap_or(false);
            ui.add_enabled_ui(f.rule.is_some(), |ui| {
                let label = if valid { format!("{}  Save", icon::FLOPPY_DISK) } else { format!("{}  Save anyway", icon::FLOPPY_DISK) };
                if primary_button(ui, &label).clicked() {
                    outcome = Outcome::Save;
                }
            });
            if secondary_button(ui, "Cancel").clicked() {
                outcome = Outcome::Close;
            }
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                let exists = self.project.as_ref().map(|p| DetectionRule::load(p).is_some()).unwrap_or(false);
                if exists && ui.link(it(format!("{}  Remove rule", icon::TRASH)).color(p.danger)).clicked() {
                    outcome = Outcome::Remove;
                }
            });
        });
        outcome
    }
}
