//! Custom-painted controls: ribbon buttons, groups, chips, segmented controls.

use eframe::egui::{
    self, Align2, Color32, CornerRadius, FontId, Rect, Response, Sense, Stroke, StrokeKind, Ui, Vec2,
};

use crate::theme::{icon_font, semibold, Palette};

fn is_icon_char(c: char) -> bool {
    ('\u{E000}'..='\u{F8FF}').contains(&c)
}

/// Build a layout job where icon glyphs use the Phosphor family and everything
/// else uses `font`.
pub fn icon_job(text: &str, font: FontId, color: Color32) -> egui::text::LayoutJob {
    let mut job = egui::text::LayoutJob::default();
    let mut buf = String::new();
    let mut buf_icon = false;
    let flush = |job: &mut egui::text::LayoutJob, buf: &mut String, icon: bool| {
        if buf.is_empty() {
            return;
        }
        let f = if icon { icon_font(font.size * 1.08) } else { font.clone() };
        job.append(buf, 0.0, egui::TextFormat { font_id: f, color, valign: egui::Align::Center, ..Default::default() });
        buf.clear();
    };
    for c in text.chars() {
        let ic = is_icon_char(c);
        if ic != buf_icon {
            flush(&mut job, &mut buf, buf_icon);
            buf_icon = ic;
        }
        buf.push(c);
    }
    flush(&mut job, &mut buf, buf_icon);
    job
}

/// RichText-like builder that renders embedded Phosphor icons correctly.
#[derive(Clone)]
pub struct IconText {
    text: String,
    size: Option<f32>,
    color: Option<Color32>,
    font: Option<FontId>,
}

pub fn it(text: impl Into<String>) -> IconText {
    IconText { text: text.into(), size: None, color: None, font: None }
}

impl IconText {
    pub fn size(mut self, s: f32) -> Self {
        self.size = Some(s);
        self
    }
    pub fn color(mut self, c: Color32) -> Self {
        self.color = Some(c);
        self
    }
    pub fn font(mut self, f: FontId) -> Self {
        self.font = Some(f);
        self
    }
    pub fn monospace(self) -> Self {
        let size = self.size.unwrap_or(13.0);
        self.font(FontId::monospace(size))
    }
    pub fn strong(self) -> Self {
        let size = self.size.unwrap_or(13.5);
        self.font(semibold(size))
    }
}

impl From<IconText> for egui::WidgetText {
    fn from(t: IconText) -> Self {
        let mut font = t.font.unwrap_or_else(|| FontId::proportional(13.5));
        if let Some(s) = t.size {
            font.size = s;
        }
        let color = t.color.unwrap_or(Color32::PLACEHOLDER);
        egui::WidgetText::LayoutJob(std::sync::Arc::new(icon_job(&t.text, font, color)))
    }
}



pub fn palette(ui: &Ui) -> Palette {
    ui.ctx().data(|d| d.get_temp::<Palette>(egui::Id::new("palette"))).unwrap_or_else(Palette::light)
}

/// Large ribbon button: icon above a (up to two-line) label.
pub fn ribbon_large(ui: &mut Ui, icon: &str, label: &str, enabled: bool, accent: bool) -> Response {
    let p = palette(ui);
    let label_galley_w = {
        let g = ui.painter().layout_no_wrap(label.to_string(), FontId::proportional(12.0), p.text);
        g.size().x
    };
    let width = (label_galley_w + 16.0).clamp(54.0, 92.0);
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(width, 70.0), Sense::click());
    let resp = if enabled { resp } else { resp.on_hover_cursor(egui::CursorIcon::NotAllowed) };
    if ui.is_rect_visible(rect) {
        let painter = ui.painter();
        let hovered = enabled && resp.hovered();
        let pressed = enabled && resp.is_pointer_button_down_on();
        if pressed {
            painter.rect_filled(rect, CornerRadius::same(6), p.pressed);
        } else if hovered {
            painter.rect_filled(rect, CornerRadius::same(6), p.hover);
        }
        let mut icon_color = if accent { p.accent } else { p.text };
        let mut text_color = p.text;
        if !enabled {
            icon_color = p.text_faint.gamma_multiply(0.7);
            text_color = p.text_faint;
        }
        if accent && enabled {
            let badge = Rect::from_center_size(egui::pos2(rect.center().x, rect.top() + 22.0), Vec2::splat(34.0));
            painter.rect_filled(badge, CornerRadius::same(8), if hovered { p.accent_hover } else { p.accent });
            icon_color = p.on_accent;
        }
        painter.text(
            egui::pos2(rect.center().x, rect.top() + 22.0),
            Align2::CENTER_CENTER,
            icon,
            icon_font(24.0),
            icon_color,
        );
        let mut job = icon_job(label, FontId::proportional(12.0), text_color);
        job.wrap.max_width = width - 6.0;
        job.halign = egui::Align::Center;
        let galley = painter.layout_job(job);
        let pos = egui::pos2(rect.center().x, rect.top() + 44.0);
        painter.galley(pos, galley, text_color);
    }
    if enabled {
        resp.on_hover_cursor(egui::CursorIcon::PointingHand)
    } else {
        resp
    }
}

/// Small ribbon button: icon + label in one row (stacked three per column).
pub fn ribbon_small(ui: &mut Ui, icon: &str, label: &str, enabled: bool) -> Response {
    let p = palette(ui);
    let font = FontId::proportional(12.5);
    let galley = ui.painter().layout_no_wrap(label.to_string(), font.clone(), p.text);
    let width = galley.size().x + 34.0;
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(width, 22.0), Sense::click());
    if ui.is_rect_visible(rect) {
        let painter = ui.painter();
        let hovered = enabled && resp.hovered();
        if enabled && resp.is_pointer_button_down_on() {
            painter.rect_filled(rect, CornerRadius::same(4), p.pressed);
        } else if hovered {
            painter.rect_filled(rect, CornerRadius::same(4), p.hover);
        }
        let (ic, tc) = if enabled { (p.text_muted, p.text) } else { (p.text_faint, p.text_faint) };
        painter.text(
            egui::pos2(rect.left() + 13.0, rect.center().y),
            Align2::CENTER_CENTER,
            icon,
            icon_font(15.0),
            ic,
        );
        painter.text(egui::pos2(rect.left() + 26.0, rect.center().y), Align2::LEFT_CENTER, label, font, tc);
    }
    resp
}

/// A small ribbon checkbox-like toggle.
pub fn ribbon_toggle(ui: &mut Ui, value: &mut bool, label: &str) -> Response {
    let p = palette(ui);
    let font = FontId::proportional(12.5);
    let galley = ui.painter().layout_no_wrap(label.to_string(), font.clone(), p.text);
    let (rect, mut resp) = ui.allocate_exact_size(Vec2::new(galley.size().x + 50.0, 22.0), Sense::click());
    if resp.clicked() {
        *value = !*value;
        resp.mark_changed();
    }
    if ui.is_rect_visible(rect) {
        let painter = ui.painter();
        if resp.hovered() {
            painter.rect_filled(rect, CornerRadius::same(4), p.hover);
        }
        let track = Rect::from_min_size(egui::pos2(rect.left() + 4.0, rect.center().y - 7.0), Vec2::new(28.0, 14.0));
        let t = ui.ctx().animate_bool(resp.id, *value);
        let bg = if *value { p.accent } else { p.border_strong };
        painter.rect_filled(track, CornerRadius::same(7), bg);
        let knob_x = egui::lerp(track.left() + 7.0..=track.right() - 7.0, t);
        painter.circle_filled(egui::pos2(knob_x, track.center().y), 5.0, Color32::WHITE);
        painter.text(egui::pos2(rect.left() + 38.0, rect.center().y), Align2::LEFT_CENTER, label, font, p.text);
    }
    resp
}

/// A ribbon group: content on top, caption at the bottom, separator on the right.
pub fn ribbon_group<R>(ui: &mut Ui, caption: &str, add: impl FnOnce(&mut Ui) -> R) -> R {
    let p = palette(ui);
    let r = ui
        .vertical(|ui| {
            ui.set_min_height(92.0);
            let inner = ui
                .horizontal_top(|ui| {
                    ui.spacing_mut().item_spacing.x = 2.0;
                    ui.set_min_height(72.0);
                    add(ui)
                })
                .inner;
            ui.add_space(1.0);
            let w = ui.min_rect().width();
            let (rect, _) = ui.allocate_exact_size(Vec2::new(w, 14.0), Sense::hover());
            ui.painter().text(rect.center(), Align2::CENTER_CENTER, caption, FontId::proportional(11.0), p.text_faint);
            inner
        })
        .inner;
    ui.add_space(6.0);
    let top = ui.min_rect().top() + 6.0;
    let x = ui.cursor().left();
    ui.painter().line_segment([egui::pos2(x, top), egui::pos2(x, top + 80.0)], Stroke::new(1.0, p.border));
    ui.add_space(8.0);
    r
}

/// A column of up to three small ribbon buttons.
pub fn ribbon_stack<R>(ui: &mut Ui, add: impl FnOnce(&mut Ui) -> R) -> R {
    ui.vertical(|ui| {
        ui.spacing_mut().item_spacing.y = 2.0;
        ui.add_space(2.0);
        add(ui)
    })
    .inner
}

/// Rounded chip / badge.
pub fn chip(ui: &mut Ui, text: &str, fg: Color32, bg: Color32) -> Response {
    let galley = ui.painter().layout_job(icon_job(text, FontId::proportional(11.5), fg));
    let size = galley.size() + Vec2::new(14.0, 6.0);
    let (rect, resp) = ui.allocate_exact_size(size, Sense::hover());
    if ui.is_rect_visible(rect) {
        ui.painter().rect_filled(rect, CornerRadius::same(10), bg);
        ui.painter().galley(rect.center() - galley.size() / 2.0, galley, fg);
    }
    resp
}

/// Mono chip for a command name.
pub fn code_chip(ui: &mut Ui, text: &str) -> Response {
    let p = palette(ui);
    let font = FontId::monospace(11.5);
    let galley = ui.painter().layout_no_wrap(text.to_string(), font, p.text_muted);
    let size = galley.size() + Vec2::new(12.0, 6.0);
    let (rect, resp) = ui.allocate_exact_size(size, Sense::hover());
    if ui.is_rect_visible(rect) {
        ui.painter().rect(rect, CornerRadius::same(4), p.hover, Stroke::NONE, StrokeKind::Inside);
        ui.painter().galley(rect.center() - galley.size() / 2.0, galley, p.text_muted);
    }
    resp
}

/// Segmented control (like a Fluent "TabList" with pill selection).
pub fn segmented<T: PartialEq + Copy>(ui: &mut Ui, value: &mut T, items: &[(T, String)]) -> bool {
    let p = palette(ui);
    let font = semibold(13.0);
    let widths: Vec<f32> = items
        .iter()
        .map(|(_, l)| ui.painter().layout_no_wrap(l.clone(), font.clone(), p.text).size().x + 28.0)
        .collect();
    let total: f32 = widths.iter().sum::<f32>() + 6.0;
    let (outer, _) = ui.allocate_exact_size(Vec2::new(total, 34.0), Sense::hover());
    ui.painter().rect_filled(outer, CornerRadius::same(8), p.hover);
    let mut x = outer.left() + 3.0;
    let mut changed = false;
    for ((item, label), w) in items.iter().zip(widths) {
        let r = Rect::from_min_size(egui::pos2(x, outer.top() + 3.0), Vec2::new(w, 28.0));
        let resp = ui.interact(r, ui.id().with(("seg", label.as_str())), Sense::click());
        let selected = *value == *item;
        if selected {
            ui.painter().rect(r, CornerRadius::same(6), p.card_bg, Stroke::new(1.0, p.border), StrokeKind::Inside);
        } else if resp.hovered() {
            ui.painter().rect_filled(r, CornerRadius::same(6), p.pressed.gamma_multiply(0.6));
        }
        let color = if selected { p.text } else { p.text_muted };
        ui.painter().text(r.center(), Align2::CENTER_CENTER, label, font.clone(), color);
        if resp.clicked() && !selected {
            *value = *item;
            changed = true;
        }
        x += w;
    }
    changed
}

/// Primary (accent filled) button.
pub fn primary_button(ui: &mut Ui, text: &str) -> Response {
    let p = palette(ui);
    ui.add(
        egui::Button::new(it(text).color(p.on_accent).font(semibold(13.5)))
            .fill(p.accent)
            .stroke(Stroke::NONE)
            .corner_radius(CornerRadius::same(6))
            .min_size(Vec2::new(0.0, 30.0)),
    )
}

pub fn secondary_button(ui: &mut Ui, text: &str) -> Response {
    ui.add(egui::Button::new(it(text)).corner_radius(CornerRadius::same(6)).min_size(Vec2::new(0.0, 30.0)))
}

/// Field caption used in forms.
pub fn field_label(ui: &mut Ui, text: &str, hint: Option<&str>) {
    let p = palette(ui);
    ui.horizontal(|ui| {
        ui.label(egui::RichText::new(text).font(semibold(12.5)).color(p.text));
        if let Some(h) = hint {
            ui.label(egui::RichText::new(h).size(11.5).color(p.text_faint));
        }
    });
}

pub fn muted(ui: &mut Ui, text: &str) -> Response {
    let p = palette(ui);
    ui.label(egui::RichText::new(text).size(12.5).color(p.text_muted))
}

/// Icon-only flat button with tooltip.
pub fn icon_button(ui: &mut Ui, icon: &str, tooltip: &str) -> Response {
    let p = palette(ui);
    let (rect, resp) = ui.allocate_exact_size(Vec2::splat(28.0), Sense::click());
    if ui.is_rect_visible(rect) {
        if resp.hovered() {
            ui.painter().rect_filled(rect, CornerRadius::same(6), p.hover);
        }
        let color = if ui.is_enabled() { p.text_muted } else { p.text_faint.gamma_multiply(0.6) };
        ui.painter().text(rect.center(), Align2::CENTER_CENTER, icon, icon_font(16.0), color);
    }
    resp.on_hover_text(tooltip)
}
