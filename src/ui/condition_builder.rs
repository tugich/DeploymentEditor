//! Visual builder for "Run only if (condition)" — shown under the text field in
//! the Options tab, modelled on the condition tab of a ConfigMgr task sequence.

use eframe::egui::{self, CornerRadius, FontId, Sense, Stroke, Vec2};
use egui_phosphor::regular as icon;

use crate::app::DeploymentEditorApp;
use crate::condition::{
    self, Cond, CondKind, Group, GroupMode, Node, Op, SystemFlag, DEPLOY_MODES, KNOWN_VARIABLES, PATH_VARIABLES, WINDOWS_VERSIONS,
};
use crate::project::Action;
use crate::theme::semibold;
use crate::widgets::{icon_button, it, muted, segmented};

/// Builder state for the selected action (kept between frames).
pub struct CondUi {
    pub action_id: Option<i64>,
    /// The condition text the tree was built from / last written.
    pub source: String,
    pub root: Group,
    /// Selected node; empty = the root group.
    pub selected: Vec<usize>,
    pub open: bool,
}

impl Default for CondUi {
    fn default() -> Self {
        Self { action_id: None, source: String::new(), root: Group::default(), selected: Vec::new(), open: true }
    }
}

fn kind_icon(k: CondKind) -> &'static str {
    match k {
        CondKind::Variable => icon::CURRENCY_DOLLAR_SIMPLE,
        CondKind::System => icon::DESKTOP_TOWER,
        CondKind::OsVersion => icon::WINDOWS_LOGO,
        CondKind::DeployMode => icon::EYE,
        CondKind::File => icon::FILE,
        CondKind::RegistryKey => icon::TREE_STRUCTURE,
        CondKind::RegistryValue => icon::LIST_DASHES,
        CondKind::App => icon::PACKAGE,
        CondKind::Process => icon::GEAR_SIX,
        CondKind::Expression => icon::TERMINAL,
    }
}

fn group_text(g: &Group, root: bool) -> String {
    let what = match g.mode {
        GroupMode::All => "all",
        GroupMode::Any => "any",
        GroupMode::None => "none",
    };
    if root {
        format!("Run this step if {what} of these are true:")
    } else {
        format!("If {what} of these are true")
    }
}

/// One clickable, full-width row of the tree.
fn tree_row(ui: &mut egui::Ui, depth: usize, glyph: &str, text: &str, selected: bool, bold: bool, problem: bool) -> egui::Response {
    let p = crate::widgets::palette(ui);
    let w = ui.available_width();
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(w, 24.0), Sense::click());
    if selected {
        ui.painter().rect_filled(rect, CornerRadius::same(5), p.accent_soft);
    } else if resp.hovered() {
        ui.painter().rect_filled(rect, CornerRadius::same(5), p.hover);
    }
    let x = rect.left() + 6.0 + depth as f32 * 18.0;
    if depth > 0 {
        // tree guide
        let gx = x - 10.0;
        ui.painter().line_segment(
            [egui::pos2(gx, rect.top()), egui::pos2(gx, rect.center().y)],
            Stroke::new(1.0, p.border_strong),
        );
        ui.painter().line_segment(
            [egui::pos2(gx, rect.center().y), egui::pos2(gx + 6.0, rect.center().y)],
            Stroke::new(1.0, p.border_strong),
        );
    }
    let icon_col = if problem { p.warning } else if selected { p.accent_text } else { p.accent };
    ui.painter().text(egui::pos2(x, rect.center().y), egui::Align2::LEFT_CENTER, glyph, crate::theme::icon_font(14.0), icon_col);
    let font = if bold { semibold(12.5) } else { FontId::proportional(12.5) };
    let max_w = (rect.right() - (x + 20.0) - 6.0).max(20.0);
    let mut job = egui::text::LayoutJob::simple_singleline(text.to_string(), font, if selected { p.accent_text } else { p.text });
    job.wrap = egui::text::TextWrapping::truncate_at_width(max_w);
    let galley = ui.painter().layout_job(job);
    ui.painter().galley(egui::pos2(x + 20.0, rect.center().y - galley.size().y / 2.0), galley, p.text);
    resp
}

fn text_field(ui: &mut egui::Ui, value: &mut String, hint: &str, mono: bool, width: f32) -> bool {
    let mut te = egui::TextEdit::singleline(value).hint_text(hint).desired_width(width).margin(egui::Margin::symmetric(7, 5));
    if mono {
        te = te.font(egui::TextStyle::Monospace);
    }
    ui.add(te).changed()
}

/// Text field with a ▾ menu of suggestions to insert (append) or set.
fn field_with_menu(ui: &mut egui::Ui, value: &mut String, hint: &str, items: &[(String, String)], append: bool) -> bool {
    let mut changed = false;
    ui.horizontal(|ui| {
        ui.spacing_mut().item_spacing.x = 4.0;
        changed |= text_field(ui, value, hint, true, ui.available_width() - 30.0);
        ui.menu_button(it(icon::CARET_DOWN).size(13.0), |ui| {
            ui.set_min_width(300.0);
            egui::ScrollArea::vertical().max_height(320.0).show(ui, |ui| {
                for (v, desc) in items {
                    let r = ui.add(egui::Button::new(egui::RichText::new(v).monospace()).frame(false)).on_hover_text(desc);
                    if r.clicked() {
                        if append {
                            let base = value.trim_end_matches('\\').to_string();
                            *value = if base.is_empty() { format!("{v}\\") } else { format!("{v}\\{}", base.trim_start_matches('\\')) };
                        } else {
                            *value = v.clone();
                        }
                        changed = true;
                        ui.close();
                    }
                    if !desc.is_empty() {
                        ui.label(egui::RichText::new(desc).size(11.0).color(crate::widgets::palette(ui).text_faint));
                    }
                    ui.add_space(2.0);
                }
            });
        });
    });
    changed
}

fn op_combo(ui: &mut egui::Ui, id: &str, op: &mut Op, ops: &[Op]) -> bool {
    let before = *op;
    egui::ComboBox::from_id_salt(id).selected_text(op.label()).width(ui.available_width()).show_ui(ui, |ui| {
        for o in ops {
            ui.selectable_value(op, *o, o.label());
        }
    });
    *op != before
}

fn small_label(ui: &mut egui::Ui, text: &str) {
    ui.add_space(2.0);
    ui.label(egui::RichText::new(text).size(11.5).color(crate::widgets::palette(ui).text_muted));
}

impl DeploymentEditorApp {
    /// Draws the builder; returns true when it changed `d.conditions`.
    pub(super) fn condition_builder(&mut self, ui: &mut egui::Ui, d: &mut Action) -> bool {
        let p = self.palette;
        // Keep the tree in sync with the text (other step selected, typed text, undo …)
        if self.cond_ui.action_id != Some(d.id) {
            self.cond_ui.action_id = Some(d.id);
            self.cond_ui.source = d.conditions.clone();
            self.cond_ui.root = condition::parse(&d.conditions);
            self.cond_ui.selected.clear();
        } else if self.cond_ui.source != d.conditions {
            self.cond_ui.source = d.conditions.clone();
            self.cond_ui.root = condition::parse(&d.conditions);
            if self.cond_ui.root.node(&self.cond_ui.selected).is_none() {
                self.cond_ui.selected.clear();
            }
        }

        let mut edited = false;
        let header = ui.horizontal(|ui| {
            let caret = if self.cond_ui.open { icon::CARET_DOWN } else { icon::CARET_RIGHT };
            ui.add(egui::Button::new(it(format!("{caret}  {}  Condition builder", icon::TREE_VIEW)).size(12.5).color(p.accent_text)).frame(false))
        });
        if header.inner.clicked() {
            self.cond_ui.open = !self.cond_ui.open;
        }
        if !self.cond_ui.open {
            return false;
        }

        egui::Frame::new()
            .fill(p.card_bg)
            .stroke(Stroke::new(1.0, p.border))
            .corner_radius(CornerRadius::same(8))
            .inner_margin(egui::Margin::same(8))
            .show(ui, |ui| {
                ui.set_width(ui.available_width());
                // Toolbar --------------------------------------------------
                ui.horizontal_wrapped(|ui| {
                    ui.spacing_mut().item_spacing.x = 2.0;
                    ui.menu_button(it(format!("{}  Condition {}", icon::PLUS_CIRCLE, icon::CARET_DOWN)).size(12.5), |ui| {
                        ui.set_min_width(220.0);
                        for k in CondKind::ALL {
                            if ui.add(egui::Button::new(it(format!("{}  {}", kind_icon(k), k.label())).size(13.0)).frame(false)).clicked() {
                                let sel = self.cond_ui.selected.clone();
                                self.cond_ui.selected = self.cond_ui.root.insert(&sel, Node::Leaf(k.new_cond()));
                                edited = true;
                                ui.close();
                            }
                        }
                    });
                    ui.menu_button(it(format!("{}  Group {}", icon::BRACKETS_CURLY, icon::CARET_DOWN)).size(12.5), |ui| {
                        ui.set_min_width(220.0);
                        for (mode, text) in [
                            (GroupMode::All, "All conditions are true"),
                            (GroupMode::Any, "Any condition is true"),
                            (GroupMode::None, "None of the conditions are true"),
                        ] {
                            if ui.add(egui::Button::new(text).frame(false)).clicked() {
                                let sel = self.cond_ui.selected.clone();
                                self.cond_ui.selected = self.cond_ui.root.insert(&sel, Node::Group(Group { mode, children: Vec::new() }));
                                edited = true;
                                ui.close();
                            }
                        }
                    });
                    let has_sel = !self.cond_ui.selected.is_empty();
                    ui.add_enabled_ui(has_sel, |ui| {
                        if icon_button(ui, icon::ARROW_UP, "Move up").clicked() {
                            let s = self.cond_ui.selected.clone();
                            self.cond_ui.selected = self.cond_ui.root.move_by(&s, -1);
                            edited = true;
                        }
                        if icon_button(ui, icon::ARROW_DOWN, "Move down").clicked() {
                            let s = self.cond_ui.selected.clone();
                            self.cond_ui.selected = self.cond_ui.root.move_by(&s, 1);
                            edited = true;
                        }
                        if icon_button(ui, icon::X, "Remove").clicked() {
                            let s = self.cond_ui.selected.clone();
                            self.cond_ui.selected = self.cond_ui.root.remove(&s);
                            edited = true;
                        }
                    });
                    ui.add_enabled_ui(!self.cond_ui.root.children.is_empty(), |ui| {
                        if icon_button(ui, icon::TRASH, "Remove all").clicked() {
                            self.cond_ui.root = Group::default();
                            self.cond_ui.selected.clear();
                            edited = true;
                        }
                    });
                });
                ui.add_space(4.0);

                // Tree ---------------------------------------------------
                egui::Frame::new()
                    .fill(p.input_bg)
                    .stroke(Stroke::new(1.0, p.border))
                    .corner_radius(CornerRadius::same(6))
                    .inner_margin(egui::Margin::same(4))
                    .show(ui, |ui| {
                        ui.set_width(ui.available_width());
                        ui.spacing_mut().item_spacing.y = 1.0;
                        let root_sel = self.cond_ui.selected.is_empty();
                        let text = group_text(&self.cond_ui.root, true);
                        if tree_row(ui, 0, icon::FUNNEL, &text, root_sel, true, false).clicked() {
                            self.cond_ui.selected.clear();
                        }
                        if self.cond_ui.root.children.is_empty() {
                            ui.add_space(2.0);
                            ui.horizontal(|ui| {
                                ui.add_space(26.0);
                                ui.label(egui::RichText::new("No conditions — the step always runs.").size(12.0).color(p.text_faint).italics());
                            });
                            ui.add_space(2.0);
                        }
                        let mut clicked: Option<Vec<usize>> = None;
                        draw_children(ui, &self.cond_ui.root, &mut Vec::new(), &self.cond_ui.selected, &mut clicked);
                        if let Some(c) = clicked {
                            self.cond_ui.selected = c;
                        }
                    });
                ui.add_space(6.0);

                // Properties of the selection ----------------------------
                let sel = self.cond_ui.selected.clone();
                let project_vars: Vec<(String, String)> = {
                    let mut v: Vec<(String, String)> = self
                        .actions
                        .iter()
                        .filter(|a| !a.variable_name.trim().is_empty())
                        .map(|a| (a.variable_name.trim().to_string(), format!("Result of step \"{}\"", a.name)))
                        .collect();
                    v.dedup_by(|a, b| a.0 == b.0);
                    v.extend(KNOWN_VARIABLES.iter().map(|(n, d)| (n.to_string(), d.to_string())));
                    v
                };
                let id = format!("cond-{}-{:?}", d.id, sel);
                if sel.is_empty() {
                    edited |= group_editor(ui, &mut self.cond_ui.root, &id);
                } else {
                    match self.cond_ui.root.node_mut(&sel) {
                        Some(Node::Group(g)) => edited |= group_editor(ui, g, &id),
                        Some(Node::Leaf(c)) => edited |= leaf_editor(ui, c, &id, &project_vars),
                        None => {}
                    }
                }
            });

        if edited {
            d.conditions = self.cond_ui.root.to_ps();
            self.cond_ui.source = d.conditions.clone();
        }
        edited
    }
}

fn draw_children(ui: &mut egui::Ui, g: &Group, path: &mut Vec<usize>, selected: &[usize], clicked: &mut Option<Vec<usize>>) {
    for (i, child) in g.children.iter().enumerate() {
        path.push(i);
        let depth = path.len();
        let is_sel = path.as_slice() == selected;
        match child {
            Node::Leaf(c) => {
                let problem = !c.problems().is_empty();
                let r = tree_row(ui, depth, kind_icon(c.kind()), &c.describe(), is_sel, false, problem);
                let r = if problem { r.on_hover_text(c.problems().join("\n")) } else { r.on_hover_text(c.describe()) };
                if r.clicked() {
                    *clicked = Some(path.clone());
                }
            }
            Node::Group(sub) => {
                if tree_row(ui, depth, icon::BRACKETS_CURLY, &group_text(sub, false), is_sel, true, sub.children.is_empty()).clicked() {
                    *clicked = Some(path.clone());
                }
                draw_children(ui, sub, path, selected, clicked);
            }
        }
        path.pop();
    }
}

fn group_editor(ui: &mut egui::Ui, g: &mut Group, id: &str) -> bool {
    let _ = id;
    small_label(ui, "The step runs if");
    let mut mode = g.mode;
    let changed = segmented(
        ui,
        &mut mode,
        &[(GroupMode::All, "All are true".into()), (GroupMode::Any, "Any is true".into()), (GroupMode::None, "None is true".into())],
    );
    g.mode = mode;
    if g.children.is_empty() {
        muted(ui, "Add conditions with \"Condition\" — they are added to the selected group.");
    }
    changed
}

fn leaf_editor(ui: &mut egui::Ui, c: &mut Cond, id: &str, vars: &[(String, String)]) -> bool {
    let p = crate::widgets::palette(ui);
    let mut changed = false;
    // Type
    let mut kind = c.kind();
    egui::ComboBox::from_id_salt(format!("{id}-kind"))
        .selected_text(it(format!("{}  {}", kind_icon(kind), kind.label())).size(13.0))
        .width(ui.available_width())
        .show_ui(ui, |ui| {
            for k in CondKind::ALL {
                ui.selectable_value(&mut kind, k, it(format!("{}  {}", kind_icon(k), k.label())).size(13.0));
            }
        });
    if kind != c.kind() {
        *c = kind.new_cond();
        changed = true;
    }
    let paths: Vec<(String, String)> = PATH_VARIABLES.iter().map(|(a, b)| (a.to_string(), b.to_string())).collect();

    match c {
        Cond::Variable { name, op, value } => {
            small_label(ui, "Variable");
            changed |= field_with_menu(ui, name, "$variable", vars, false);
            small_label(ui, "Condition");
            changed |= op_combo(ui, &format!("{id}-op"), op, &Op::ALL);
            small_label(ui, "Value");
            changed |= text_field(ui, value, "text, number, $true, $false or $null", true, f32::INFINITY);
        }
        Cond::System { flag, is } => {
            small_label(ui, "Condition");
            let current = if *is { flag.texts().0 } else { flag.texts().1 };
            egui::ComboBox::from_id_salt(format!("{id}-sys")).selected_text(current).width(ui.available_width()).show_ui(ui, |ui| {
                for f in SystemFlag::ALL {
                    for b in [true, false] {
                        let t = if b { f.texts().0 } else { f.texts().1 };
                        if ui.selectable_label(*flag == f && *is == b, t).clicked() {
                            *flag = f;
                            *is = b;
                            changed = true;
                        }
                    }
                }
            });
            ui.label(egui::RichText::new(format!("Uses the PSADT variable {}", flag.var())).size(11.0).color(p.text_faint));
        }
        Cond::OsVersion { op, version } => {
            small_label(ui, "Windows version is");
            changed |= op_combo(ui, &format!("{id}-op"), op, &Op::ORDER);
            let items: Vec<(String, String)> = WINDOWS_VERSIONS.iter().map(|(v, n)| (v.to_string(), n.to_string())).collect();
            changed |= field_with_menu(ui, version, "10.0.22631", &items, false);
        }
        Cond::DeployMode { mode, is } => {
            small_label(ui, "Deploy mode");
            changed |= segmented(ui, is, &[(true, "is".into()), (false, "is not".into())]);
            egui::ComboBox::from_id_salt(format!("{id}-dm")).selected_text(mode.as_str()).width(ui.available_width()).show_ui(ui, |ui| {
                for m in DEPLOY_MODES {
                    if ui.selectable_label(mode == m, m).clicked() {
                        *mode = m.to_string();
                        changed = true;
                    }
                }
            });
        }
        Cond::File { path, folder, exists } => {
            changed |= segmented(ui, folder, &[(false, "File".into()), (true, "Folder".into())]);
            small_label(ui, "Path");
            changed |= field_with_menu(ui, path, "$envProgramFiles\\Vendor\\App.exe", &paths, true);
            changed |= segmented(ui, exists, &[(true, "exists".into()), (false, "does not exist".into())]);
        }
        Cond::RegistryKey { key, exists } => {
            small_label(ui, "Key");
            changed |= text_field(ui, key, "HKLM\\SOFTWARE\\Vendor\\App", true, f32::INFINITY);
            changed |= segmented(ui, exists, &[(true, "exists".into()), (false, "does not exist".into())]);
        }
        Cond::RegistryValue { key, name, op, data, exists } => {
            small_label(ui, "Key");
            changed |= text_field(ui, key, "HKLM\\SOFTWARE\\Vendor\\App", true, f32::INFINITY);
            small_label(ui, "Value name");
            changed |= text_field(ui, name, "Version", true, f32::INFINITY);
            small_label(ui, "Condition");
            let current = match op {
                None if *exists => "exists".to_string(),
                None => "does not exist".to_string(),
                Some(o) => o.label().to_string(),
            };
            egui::ComboBox::from_id_salt(format!("{id}-rv")).selected_text(current).width(ui.available_width()).show_ui(ui, |ui| {
                if ui.selectable_label(op.is_none() && *exists, "exists").clicked() {
                    *op = None;
                    *exists = true;
                    changed = true;
                }
                if ui.selectable_label(op.is_none() && !*exists, "does not exist").clicked() {
                    *op = None;
                    *exists = false;
                    changed = true;
                }
                for o in Op::ALL {
                    if ui.selectable_label(*op == Some(o), o.label()).clicked() {
                        *op = Some(o);
                        *exists = true;
                        changed = true;
                    }
                }
            });
            if op.is_some() {
                small_label(ui, "Data");
                changed |= text_field(ui, data, "1.2.3", true, f32::INFINITY);
            }
        }
        Cond::App { name, exact, installed } => {
            small_label(ui, "Application name (as in Apps & features)");
            changed |= text_field(ui, name, "7-Zip", false, f32::INFINITY);
            changed |= ui.checkbox(exact, "Exact name (otherwise: name contains)").changed();
            changed |= segmented(ui, installed, &[(true, "is installed".into()), (false, "is not installed".into())]);
        }
        Cond::Process { name, running } => {
            small_label(ui, "Process name");
            changed |= text_field(ui, name, "excel", false, f32::INFINITY);
            changed |= segmented(ui, running, &[(true, "is running".into()), (false, "is not running".into())]);
        }
        Cond::Expression { code } => {
            small_label(ui, "PowerShell expression (true = run the step)");
            changed |= ui
                .add(egui::TextEdit::multiline(code).font(egui::TextStyle::Monospace).desired_rows(3).desired_width(f32::INFINITY))
                .changed();
        }
    }
    for prob in c.problems() {
        ui.label(it(format!("{}  {prob}", icon::WARNING)).size(11.5).color(p.warning));
    }
    changed
}
