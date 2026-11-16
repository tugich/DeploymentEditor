//! Colours, fonts and egui style for a Fluent / Office-like look.

use eframe::egui::{
    self, epaint::Shadow, Color32, CornerRadius, FontData, FontDefinitions, FontFamily, FontId, Margin,
    Stroke, TextStyle, Visuals,
};
use std::sync::Arc;

pub const SEMIBOLD: &str = "semibold";

/// Brand colour of the Deployment Editor logo.
pub const BRAND_BLUE: Color32 = Color32::from_rgb(0x15, 0x7E, 0xFB);

#[derive(Clone, Copy, Debug)]
pub struct Palette {
    pub dark: bool,
    pub accent: Color32,
    pub accent_hover: Color32,
    pub accent_soft: Color32,
    pub accent_text: Color32,
    pub on_accent: Color32,
    pub window_bg: Color32,
    pub ribbon_bg: Color32,
    pub panel_bg: Color32,
    pub canvas_bg: Color32,
    pub card_bg: Color32,
    pub card_hover: Color32,
    pub border: Color32,
    pub border_strong: Color32,
    pub hover: Color32,
    pub pressed: Color32,
    pub text: Color32,
    pub text_muted: Color32,
    pub text_faint: Color32,
    pub success: Color32,
    pub warning: Color32,
    pub danger: Color32,
    pub info: Color32,
    pub script: Color32,
    pub input_bg: Color32,
}

impl Palette {
    pub fn light() -> Self {
        Self {
            dark: false,
            accent: BRAND_BLUE,
            accent_hover: Color32::from_rgb(0x0B, 0x6B, 0xDD),
            accent_soft: Color32::from_rgb(0xE6, 0xF0, 0xFF),
            accent_text: Color32::from_rgb(0x0B, 0x5F, 0xC6),
            on_accent: Color32::WHITE,
            window_bg: Color32::from_rgb(0xF3, 0xF4, 0xF6),
            ribbon_bg: Color32::from_rgb(0xFB, 0xFB, 0xFC),
            panel_bg: Color32::from_rgb(0xFF, 0xFF, 0xFF),
            canvas_bg: Color32::from_rgb(0xF3, 0xF4, 0xF6),
            card_bg: Color32::WHITE,
            card_hover: Color32::from_rgb(0xF7, 0xFA, 0xF9),
            border: Color32::from_rgb(0xE3, 0xE5, 0xE8),
            border_strong: Color32::from_rgb(0xC9, 0xCD, 0xD2),
            hover: Color32::from_rgb(0xEC, 0xEE, 0xF1),
            pressed: Color32::from_rgb(0xDE, 0xE2, 0xE6),
            text: Color32::from_rgb(0x1F, 0x23, 0x28),
            text_muted: Color32::from_rgb(0x5B, 0x63, 0x6D),
            text_faint: Color32::from_rgb(0x8C, 0x93, 0x9C),
            success: Color32::from_rgb(0x10, 0x7C, 0x41),
            warning: Color32::from_rgb(0xB4, 0x6A, 0x00),
            danger: Color32::from_rgb(0xC4, 0x2B, 0x1C),
            info: Color32::from_rgb(0x6B, 0x4F, 0xD8),
            script: Color32::from_rgb(0x0F, 0x76, 0x6E),
            input_bg: Color32::WHITE,
        }
    }

    pub fn dark() -> Self {
        Self {
            dark: true,
            accent: BRAND_BLUE,
            accent_hover: Color32::from_rgb(0x3A, 0x93, 0xFF),
            accent_soft: Color32::from_rgb(0x15, 0x2A, 0x45),
            accent_text: Color32::from_rgb(0x6A, 0xAE, 0xFF),
            on_accent: Color32::WHITE,
            window_bg: Color32::from_rgb(0x1B, 0x1C, 0x1F),
            ribbon_bg: Color32::from_rgb(0x25, 0x26, 0x2A),
            panel_bg: Color32::from_rgb(0x22, 0x23, 0x27),
            canvas_bg: Color32::from_rgb(0x1B, 0x1C, 0x1F),
            card_bg: Color32::from_rgb(0x2A, 0x2B, 0x30),
            card_hover: Color32::from_rgb(0x31, 0x33, 0x38),
            border: Color32::from_rgb(0x38, 0x3A, 0x40),
            border_strong: Color32::from_rgb(0x4A, 0x4D, 0x55),
            hover: Color32::from_rgb(0x33, 0x35, 0x3A),
            pressed: Color32::from_rgb(0x3D, 0x40, 0x46),
            text: Color32::from_rgb(0xEC, 0xED, 0xEF),
            text_muted: Color32::from_rgb(0xA8, 0xAE, 0xB6),
            text_faint: Color32::from_rgb(0x7A, 0x80, 0x88),
            success: Color32::from_rgb(0x4C, 0xC2, 0x7A),
            warning: Color32::from_rgb(0xF0, 0xB4, 0x4A),
            danger: Color32::from_rgb(0xF0, 0x6A, 0x5E),
            info: Color32::from_rgb(0xA9, 0x93, 0xFF),
            script: Color32::from_rgb(0x2D, 0xB3, 0xA3),
            input_bg: Color32::from_rgb(0x1E, 0x1F, 0x23),
        }
    }

    pub fn for_mode(dark: bool) -> Self {
        if dark {
            Self::dark()
        } else {
            Self::light()
        }
    }
}

fn try_load(paths: &[&str]) -> Option<Vec<u8>> {
    paths.iter().find_map(|p| std::fs::read(p).ok())
}

/// Use Segoe UI / Cascadia on Windows (falls back to Inter / DejaVu elsewhere),
/// plus the Phosphor icon font.
pub fn install_fonts(ctx: &egui::Context) {
    let mut fonts = FontDefinitions::default();
    let windir = std::env::var("WINDIR").unwrap_or_else(|_| "C:\\Windows".into());
    let wf = |f: &str| format!("{windir}\\Fonts\\{f}");

    let regular = try_load(&[
        &wf("segoeui.ttf"),
        "/usr/share/fonts/opentype/inter/Inter-Regular.otf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]);
    let semibold = try_load(&[
        &wf("seguisb.ttf"),
        &wf("segoeuib.ttf"),
        "/usr/share/fonts/opentype/inter/Inter-SemiBold.otf",
        "/usr/share/fonts/truetype/noto/NotoSans-SemiBold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]);
    let mono = try_load(&[
        &wf("CascadiaMono.ttf"),
        &wf("CascadiaCode.ttf"),
        &wf("consola.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    ]);

    if let Some(bytes) = regular {
        fonts.font_data.insert("ui-regular".into(), Arc::new(FontData::from_owned(bytes)));
        fonts.families.entry(FontFamily::Proportional).or_default().insert(0, "ui-regular".into());
    }
    if let Some(bytes) = mono {
        fonts.font_data.insert("ui-mono".into(), Arc::new(FontData::from_owned(bytes)));
        fonts.families.entry(FontFamily::Monospace).or_default().insert(0, "ui-mono".into());
    }
    // Semibold family: semibold face first, then whatever proportional has (icons, emoji).
    let mut semi_chain: Vec<String> = Vec::new();
    if let Some(bytes) = semibold {
        fonts.font_data.insert("ui-semibold".into(), Arc::new(FontData::from_owned(bytes)));
        semi_chain.push("ui-semibold".into());
    }

    egui_phosphor::add_to_fonts(&mut fonts, egui_phosphor::Variant::Regular);
    // Icons are rendered with this family only (see `icon_font`), so Phosphor must
    // come first; Inter / Segoe UI also have glyphs in the Private Use Area.
    fonts.families.insert(FontFamily::Name("phosphor".into()), vec!["phosphor".into()]);
    let prop = fonts.families.get(&FontFamily::Proportional).cloned().unwrap_or_default();
    semi_chain.extend(prop.into_iter().filter(|f| f != "ui-semibold" && f != "phosphor"));
    fonts.families.insert(FontFamily::Name(SEMIBOLD.into()), semi_chain);

    // Make sure monospace text can show icons too.
    fonts.families.entry(FontFamily::Monospace).or_default().push("phosphor".into());

    ctx.set_fonts(fonts);
}

/// Font for icon glyphs. Icons are always rendered with this explicit family because
/// some UI fonts (Inter, possibly Segoe UI) also define glyphs in the Private Use Area.
pub fn icon_font(size: f32) -> FontId {
    FontId::new(size, FontFamily::Name("phosphor".into()))
}

pub fn semibold(size: f32) -> FontId {
    FontId::new(size, FontFamily::Name(SEMIBOLD.into()))
}

pub fn apply_style(ctx: &egui::Context, p: &Palette) {
    let mut visuals = if p.dark { Visuals::dark() } else { Visuals::light() };
    visuals.panel_fill = p.panel_bg;
    visuals.window_fill = p.panel_bg;
    visuals.extreme_bg_color = p.input_bg;
    visuals.faint_bg_color = p.hover;
    visuals.code_bg_color = if p.dark { Color32::from_rgb(0x1A, 0x1B, 0x1E) } else { Color32::from_rgb(0xF6, 0xF7, 0xF9) };
    visuals.hyperlink_color = p.accent_text;
    visuals.selection.bg_fill = p.accent.gamma_multiply(if p.dark { 0.45 } else { 0.25 });
    visuals.selection.stroke = Stroke::new(1.0, p.accent);
    visuals.window_corner_radius = CornerRadius::same(10);
    visuals.menu_corner_radius = CornerRadius::same(8);
    visuals.window_stroke = Stroke::new(1.0, p.border);
    visuals.window_shadow = Shadow {
        offset: [0, 8],
        blur: 28,
        spread: 0,
        color: Color32::from_black_alpha(if p.dark { 120 } else { 40 }),
    };
    visuals.popup_shadow = Shadow {
        offset: [0, 4],
        blur: 16,
        spread: 0,
        color: Color32::from_black_alpha(if p.dark { 100 } else { 30 }),
    };
    visuals.override_text_color = Some(p.text);

    let w = &mut visuals.widgets;
    let r = CornerRadius::same(6);
    w.noninteractive.bg_stroke = Stroke::new(1.0, p.border);
    w.noninteractive.fg_stroke = Stroke::new(1.0, p.text);
    w.noninteractive.corner_radius = r;
    w.inactive.bg_fill = if p.dark { Color32::from_rgb(0x30, 0x32, 0x37) } else { Color32::from_rgb(0xF4, 0xF5, 0xF7) };
    w.inactive.weak_bg_fill = w.inactive.bg_fill;
    w.inactive.bg_stroke = Stroke::new(1.0, p.border_strong);
    w.inactive.fg_stroke = Stroke::new(1.0, p.text);
    w.inactive.corner_radius = r;
    w.hovered.bg_fill = p.hover;
    w.hovered.weak_bg_fill = p.hover;
    w.hovered.bg_stroke = Stroke::new(1.0, p.border_strong);
    w.hovered.fg_stroke = Stroke::new(1.0, p.text);
    w.hovered.corner_radius = r;
    w.hovered.expansion = 0.0;
    w.active.bg_fill = p.pressed;
    w.active.weak_bg_fill = p.pressed;
    w.active.bg_stroke = Stroke::new(1.0, p.accent);
    w.active.fg_stroke = Stroke::new(1.0, p.text);
    w.active.corner_radius = r;
    w.active.expansion = 0.0;
    w.open.bg_fill = p.hover;
    w.open.weak_bg_fill = p.hover;
    w.open.corner_radius = r;

    let theme = if p.dark { egui::Theme::Dark } else { egui::Theme::Light };
    ctx.set_theme(theme);
    ctx.set_visuals_of(theme, visuals);

    ctx.all_styles_mut(|style| {
        style.spacing.item_spacing = egui::vec2(8.0, 6.0);
        style.spacing.button_padding = egui::vec2(10.0, 5.0);
        style.spacing.interact_size.y = 28.0;
        style.spacing.menu_margin = Margin::same(6);
        style.spacing.window_margin = Margin::same(16);
        style.spacing.scroll.bar_width = 8.0;
        style.spacing.scroll.floating = true;
        style.text_styles.insert(TextStyle::Body, FontId::proportional(13.5));
        style.text_styles.insert(TextStyle::Button, FontId::proportional(13.5));
        style.text_styles.insert(TextStyle::Small, FontId::proportional(11.5));
        style.text_styles.insert(TextStyle::Monospace, FontId::monospace(13.0));
        style.text_styles.insert(TextStyle::Heading, semibold(20.0));
        style.interaction.selectable_labels = false;
    });
}
