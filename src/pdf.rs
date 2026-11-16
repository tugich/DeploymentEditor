//! A small, dependency-light PDF writer with a flowing layout (A4).
//!
//! Uses the standard PDF fonts (Helvetica, Helvetica-Bold, Helvetica-Oblique,
//! Courier) in WinAnsiEncoding, so no font files are embedded and the output
//! opens everywhere. Supports wrapped paragraphs, two-column tables, code blocks,
//! filled shapes, PNG images (with alpha), bookmarks and page footers.

use std::io::Write;

use flate2::{write::ZlibEncoder, Compression};

use crate::pdf_metrics::{HELVETICA, HELVETICA_BOLD};

pub const PAGE_W: f32 = 595.28;
pub const PAGE_H: f32 = 841.89;
pub const MARGIN_L: f32 = 52.0;
pub const MARGIN_R: f32 = 52.0;
pub const MARGIN_TOP: f32 = 54.0;
pub const MARGIN_BOTTOM: f32 = 62.0;
pub const CONTENT_W: f32 = PAGE_W - MARGIN_L - MARGIN_R;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Font {
    Regular,
    Bold,
    Italic,
    Mono,
}

impl Font {
    fn resource(&self) -> &'static str {
        match self {
            Font::Regular => "/F1",
            Font::Bold => "/F2",
            Font::Italic => "/F3",
            Font::Mono => "/F4",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Rgb(pub f32, pub f32, pub f32);

impl Rgb {
    pub const fn hex(v: u32) -> Self {
        Rgb(((v >> 16) & 0xFF) as f32 / 255.0, ((v >> 8) & 0xFF) as f32 / 255.0, (v & 0xFF) as f32 / 255.0)
    }
    fn fill(&self) -> String {
        format!("{:.3} {:.3} {:.3} rg", self.0, self.1, self.2)
    }
    fn stroke(&self) -> String {
        format!("{:.3} {:.3} {:.3} RG", self.0, self.1, self.2)
    }
}

/// Replace characters that WinAnsi cannot show with readable equivalents.
pub fn simplify(s: &str) -> String {
    s.replace('≥', ">=")
        .replace('≤', "<=")
        .replace('≠', "<>")
        .replace('→', "->")
        .replace('←', "<-")
        .replace('✓', "")
        .replace('\t', "    ")
        .replace('\u{feff}', "")
}

/// Encode a char as a WinAnsi (cp1252) byte.
pub fn winansi(c: char) -> u8 {
    let u = c as u32;
    match u {
        0x20..=0x7E | 0xA0..=0xFF => u as u8,
        0x20AC => 0x80,
        0x201A => 0x82,
        0x0192 => 0x83,
        0x201E => 0x84,
        0x2026 => 0x85,
        0x2020 => 0x86,
        0x2021 => 0x87,
        0x02C6 => 0x88,
        0x2030 => 0x89,
        0x0160 => 0x8A,
        0x2039 => 0x8B,
        0x0152 => 0x8C,
        0x017D => 0x8E,
        0x2018 => 0x91,
        0x2019 => 0x92,
        0x201C => 0x93,
        0x201D => 0x94,
        0x2022 => 0x95,
        0x2013 => 0x96,
        0x2014 => 0x97,
        0x02DC => 0x98,
        0x2122 => 0x99,
        0x0161 => 0x9A,
        0x203A => 0x9B,
        0x0153 => 0x9C,
        0x017E => 0x9E,
        0x0178 => 0x9F,
        _ => b'?',
    }
}

pub fn text_width(text: &str, font: Font, size: f32) -> f32 {
    let units: u32 = match font {
        Font::Mono => text.chars().count() as u32 * 600,
        Font::Bold => text.chars().map(|c| HELVETICA_BOLD[winansi(c) as usize] as u32).sum(),
        Font::Regular | Font::Italic => text.chars().map(|c| HELVETICA[winansi(c) as usize] as u32).sum(),
    };
    units as f32 * size / 1000.0
}

fn pdf_string(text: &str) -> String {
    let mut out = String::from("(");
    for c in text.chars() {
        let b = winansi(c);
        match b {
            b'(' | b')' | b'\\' => {
                out.push('\\');
                out.push(b as char);
            }
            0x20..=0x7E => out.push(b as char),
            _ => out.push_str(&format!("\\{b:03o}")),
        }
    }
    out.push(')');
    out
}

/// Word-wrap `text` to `width`. Long words are broken by characters.
pub fn wrap(text: &str, font: Font, size: f32, width: f32) -> Vec<String> {
    let mut lines = Vec::new();
    for raw_line in simplify(text).split('\n') {
        let raw_line = raw_line.trim_end_matches('\r');
        if raw_line.is_empty() {
            lines.push(String::new());
            continue;
        }
        let mut line = String::new();
        for word in raw_line.split(' ') {
            let candidate = if line.is_empty() { word.to_string() } else { format!("{line} {word}") };
            if text_width(&candidate, font, size) <= width {
                line = candidate;
                continue;
            }
            if !line.is_empty() {
                lines.push(std::mem::take(&mut line));
            }
            // break a word that is longer than the line
            let mut part = String::new();
            for c in word.chars() {
                let next = format!("{part}{c}");
                if text_width(&next, font, size) > width && !part.is_empty() {
                    lines.push(std::mem::take(&mut part));
                    part.push(c);
                } else {
                    part = next;
                }
            }
            line = part;
        }
        lines.push(line);
    }
    lines
}

pub struct Image {
    pub width: u32,
    pub height: u32,
    rgb: Vec<u8>,
    alpha: Option<Vec<u8>>,
}

impl Image {
    pub fn from_png(bytes: &[u8]) -> Option<Image> {
        let img = image::load_from_memory(bytes).ok()?.into_rgba8();
        let (w, h) = img.dimensions();
        let mut rgb = Vec::with_capacity((w * h * 3) as usize);
        let mut alpha = Vec::with_capacity((w * h) as usize);
        let mut has_alpha = false;
        for px in img.pixels() {
            rgb.extend_from_slice(&px.0[..3]);
            alpha.push(px.0[3]);
            has_alpha |= px.0[3] != 255;
        }
        Some(Image { width: w, height: h, rgb, alpha: has_alpha.then_some(alpha) })
    }
}

struct Bookmark {
    title: String,
    page: usize,
    y: f32,
    level: u8,
}

pub struct Pdf {
    pages: Vec<String>,
    /// Distance from the top of the current page.
    pub y: f32,
    images: Vec<Image>,
    bookmarks: Vec<Bookmark>,
    pub title: String,
    pub author: String,
    pub footer_left: String,
    /// called for every new page except the first (e.g. running header)
    pub muted: Rgb,
    pub rule: Rgb,
}

fn deflate(data: &[u8]) -> Vec<u8> {
    let mut e = ZlibEncoder::new(Vec::new(), Compression::default());
    e.write_all(data).expect("in-memory write");
    e.finish().expect("in-memory write")
}

impl Default for Pdf {
    fn default() -> Self {
        Self::new()
    }
}

impl Pdf {
    pub fn new() -> Self {
        Self {
            pages: vec![String::new()],
            y: MARGIN_TOP,
            images: Vec::new(),
            bookmarks: Vec::new(),
            title: String::new(),
            author: String::new(),
            footer_left: String::new(),
            muted: Rgb::hex(0x6B7280),
            rule: Rgb::hex(0xE3E5E8),
        }
    }

    #[cfg(test)]
    pub fn page_count(&self) -> usize {
        self.pages.len()
    }

    fn ops(&mut self) -> &mut String {
        self.pages.last_mut().expect("at least one page")
    }

    pub fn new_page(&mut self) {
        self.pages.push(String::new());
        self.y = MARGIN_TOP;
    }

    /// Start a new page if less than `h` points are left.
    pub fn ensure(&mut self, h: f32) {
        if self.y + h > PAGE_H - MARGIN_BOTTOM {
            self.new_page();
        }
    }

    pub fn space(&mut self, h: f32) {
        self.y += h;
    }

    // ---------------------------------------------------------------- primitives (top-based y)

    pub fn text_at(&mut self, x: f32, y_top: f32, font: Font, size: f32, color: Rgb, text: &str) {
        let baseline = PAGE_H - y_top - size * 0.8;
        let op = format!("BT {} {} {:.2} Tf {:.2} {:.2} Td {} Tj ET\n", color.fill(), font.resource(), size, x, baseline, pdf_string(&simplify(text)));
        self.ops().push_str(&op);
    }

    pub fn rect(&mut self, x: f32, y_top: f32, w: f32, h: f32, color: Rgb) {
        let op = format!("{} {:.2} {:.2} {:.2} {:.2} re f\n", color.fill(), x, PAGE_H - y_top - h, w, h);
        self.ops().push_str(&op);
    }

    pub fn rounded_rect(&mut self, x: f32, y_top: f32, w: f32, h: f32, r: f32, color: Rgb) {
        let (x0, y0, x1, y1) = (x, PAGE_H - y_top - h, x + w, PAGE_H - y_top);
        let k = 0.5523 * r;
        let op = format!(
            "{} {x0:.2} {:.2} m {:.2} {y0:.2} l {:.2} {y0:.2} {x1:.2} {:.2} {x1:.2} {:.2} c {x1:.2} {:.2} l {x1:.2} {:.2} {:.2} {y1:.2} {:.2} {y1:.2} c {:.2} {y1:.2} l {:.2} {y1:.2} {x0:.2} {:.2} {x0:.2} {:.2} c {x0:.2} {:.2} l {x0:.2} {:.2} {:.2} {y0:.2} {:.2} {y0:.2} c f\n",
            color.fill(),
            y0 + r,
            x1 - r,
            x1 - r + k,
            y0 + r - k,
            y0 + r,
            y1 - r,
            y1 - r + k,
            x1 - r + k,
            x1 - r,
            x0 + r,
            x0 + r - k,
            y1 - r + k,
            y1 - r,
            y0 + r,
            y0 + r - k,
            x0 + r - k,
            x0 + r,
        );
        self.ops().push_str(&op);
    }

    pub fn circle(&mut self, cx: f32, cy_top: f32, r: f32, color: Rgb) {
        self.rounded_rect(cx - r, cy_top - r, 2.0 * r, 2.0 * r, r, color);
    }

    pub fn hline(&mut self, x0: f32, x1: f32, y_top: f32, width: f32, color: Rgb) {
        let y = PAGE_H - y_top;
        let op = format!("{} {width:.2} w {x0:.2} {y:.2} m {x1:.2} {y:.2} l S\n", color.stroke());
        self.ops().push_str(&op);
    }

    /// Register a PNG; returns the image index for [`Pdf::image`].
    pub fn add_image(&mut self, img: Image) -> usize {
        self.images.push(img);
        self.images.len() - 1
    }

    pub fn image(&mut self, idx: usize, x: f32, y_top: f32, w: f32, h: f32) {
        let op = format!("q {w:.2} 0 0 {h:.2} {x:.2} {:.2} cm /Im{} Do Q\n", PAGE_H - y_top - h, idx + 1);
        self.ops().push_str(&op);
    }

    pub fn bookmark(&mut self, title: &str, level: u8) {
        let page = self.pages.len() - 1;
        self.bookmarks.push(Bookmark { title: simplify(title), page, y: self.y, level });
    }

    // ---------------------------------------------------------------- flowing layout

    /// Wrapped paragraph at the cursor, breaking pages between lines.
    pub fn paragraph(&mut self, text: &str, font: Font, size: f32, color: Rgb, x: f32, width: f32) {
        let leading = size * 1.38;
        for line in wrap(text, font, size, width) {
            self.ensure(leading);
            self.text_at(x, self.y, font, size, color, &line);
            self.y += leading;
        }
    }

    #[allow(dead_code)]
    pub fn heading(&mut self, text: &str, size: f32, color: Rgb) {
        self.ensure(size * 2.6 + 40.0);
        self.bookmark(text, 0);
        self.text_at(MARGIN_L, self.y, Font::Bold, size, color, text);
        self.y += size * 1.5;
    }

    /// Two-column key/value table with wrapped values.
    pub fn key_values(&mut self, rows: &[(String, String, Font)], key_width: f32, size: f32, text: Rgb) {
        let leading = size * 1.4;
        let value_x = MARGIN_L + key_width;
        let value_w = CONTENT_W - key_width;
        for (k, v, font) in rows {
            let lines = wrap(if v.is_empty() { "-" } else { v }, *font, size, value_w);
            let h = lines.len() as f32 * leading + 6.0;
            self.ensure(h.min(120.0));
            let top = self.y;
            let muted = self.muted;
            self.text_at(MARGIN_L, top + 3.0, Font::Regular, size, muted, k);
            for (i, l) in lines.iter().enumerate() {
                self.ensure(leading);
                if i > 0 && self.y < top {
                    // continued on a new page
                }
                self.text_at(value_x, self.y + 3.0, *font, size, text, l);
                self.y += leading;
            }
            self.y += 6.0;
            let rule = self.rule;
            self.hline(MARGIN_L, MARGIN_L + CONTENT_W, self.y - 2.0, 0.5, rule);
        }
    }

    /// Monospace code block with background, wrapped by characters.
    pub fn code_block(&mut self, code: &str, x: f32, width: f32, size: f32, bg: Rgb, fg: Rgb) {
        let leading = size * 1.35;
        let pad = 7.0;
        let max_chars = (((width - 2.0 * pad) / (size * 0.6)).floor() as usize).max(10);
        let mut lines: Vec<String> = Vec::new();
        for l in simplify(code).replace("\r\n", "\n").split('\n') {
            let chars: Vec<char> = l.chars().collect();
            if chars.is_empty() {
                lines.push(String::new());
            }
            for chunk in chars.chunks(max_chars) {
                lines.push(chunk.iter().collect());
            }
        }
        while lines.last().map(|l| l.trim().is_empty()).unwrap_or(false) {
            lines.pop();
        }
        if lines.is_empty() {
            return;
        }
        self.ensure(leading * 2.0 + pad * 2.0);
        self.rect(x, self.y, width, pad, bg);
        self.y += pad;
        for l in &lines {
            if self.y + leading > PAGE_H - MARGIN_BOTTOM {
                self.new_page();
                self.rect(x, self.y, width, pad, bg);
                self.y += pad;
            }
            self.rect(x, self.y, width, leading, bg);
            self.text_at(x + pad, self.y + (leading - size) / 2.0, Font::Mono, size, fg, l);
            self.y += leading;
        }
        self.rect(x, self.y, width, pad, bg);
        self.y += pad;
    }

    // ---------------------------------------------------------------- output

    pub fn finish(mut self) -> Vec<u8> {
        // footers
        let n = self.pages.len();
        for i in 0..n {
            let footer = format!(
                "{} {:.2} w {MARGIN_L:.2} {:.2} m {:.2} {:.2} l S\n",
                self.rule.stroke(),
                0.6,
                PAGE_H - (PAGE_H - 44.0),
                MARGIN_L + CONTENT_W,
                PAGE_H - (PAGE_H - 44.0)
            );
            let left = format!(
                "BT {} /F1 7.5 Tf {MARGIN_L:.2} 30 Td {} Tj ET\n",
                self.muted.fill(),
                pdf_string(&simplify(&self.footer_left))
            );
            let page_text = format!("Page {} of {n}", i + 1);
            let w = text_width(&page_text, Font::Regular, 7.5);
            let right = format!("BT {} /F1 7.5 Tf {:.2} 30 Td {} Tj ET\n", self.muted.fill(), MARGIN_L + CONTENT_W - w, pdf_string(&page_text));
            self.pages[i].push_str(&footer);
            self.pages[i].push_str(&left);
            self.pages[i].push_str(&right);
        }

        let mut objects: Vec<Vec<u8>> = Vec::new(); // index 0 => object 1
        let add = |data: Vec<u8>, objects: &mut Vec<Vec<u8>>| -> usize {
            objects.push(data);
            objects.len()
        };
        let stream = |dict: &str, data: &[u8]| -> Vec<u8> {
            let mut v = format!("<< {dict} /Length {} >>\nstream\n", data.len()).into_bytes();
            v.extend_from_slice(data);
            v.extend_from_slice(b"\nendstream");
            v
        };

        // reserve catalog (1) and pages (2)
        add(Vec::new(), &mut objects);
        add(Vec::new(), &mut objects);
        let fonts = [("F1", "Helvetica"), ("F2", "Helvetica-Bold"), ("F3", "Helvetica-Oblique"), ("F4", "Courier")];
        let mut font_refs = String::new();
        for (res, base) in fonts {
            let id = add(format!("<< /Type /Font /Subtype /Type1 /BaseFont /{base} /Encoding /WinAnsiEncoding >>").into_bytes(), &mut objects);
            font_refs.push_str(&format!("/{res} {id} 0 R "));
        }
        let mut image_refs = String::new();
        for (i, img) in self.images.iter().enumerate() {
            let smask = img.alpha.as_ref().map(|a| {
                let data = deflate(a);
                add(
                    stream(
                        &format!("/Type /XObject /Subtype /Image /Width {} /Height {} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode", img.width, img.height),
                        &data,
                    ),
                    &mut objects,
                )
            });
            let data = deflate(&img.rgb);
            let sm = smask.map(|id| format!(" /SMask {id} 0 R")).unwrap_or_default();
            let id = add(
                stream(
                    &format!("/Type /XObject /Subtype /Image /Width {} /Height {} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode{sm}", img.width, img.height),
                    &data,
                ),
                &mut objects,
            );
            image_refs.push_str(&format!("/Im{} {id} 0 R ", i + 1));
        }
        let resources = format!("<< /Font << {font_refs}>> /XObject << {image_refs}>> >>");

        let mut page_ids = Vec::new();
        for ops in &self.pages {
            let content = deflate(ops.as_bytes());
            let cid = add(stream("/Filter /FlateDecode", &content), &mut objects);
            // page objects reference the pages tree (2)
            let pid = add(
                format!("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_W:.2} {PAGE_H:.2}] /Resources {resources} /Contents {cid} 0 R >>").into_bytes(),
                &mut objects,
            );
            page_ids.push(pid);
        }
        let kids: String = page_ids.iter().map(|id| format!("{id} 0 R ")).collect();
        objects[1] = format!("<< /Type /Pages /Kids [{kids}] /Count {} >>", page_ids.len()).into_bytes();

        // bookmarks (flat list of top level + children)
        let mut outline_root = None;
        if !self.bookmarks.is_empty() {
            let root_id = objects.len() + 1;
            let first_item = root_id + 1;
            let count = self.bookmarks.len();
            // build a two-level tree
            let mut parents: Vec<usize> = Vec::new(); // index in bookmarks of each item's parent (usize::MAX = root)
            let mut last_top = usize::MAX;
            for (i, b) in self.bookmarks.iter().enumerate() {
                if b.level == 0 || last_top == usize::MAX {
                    parents.push(usize::MAX);
                    last_top = i;
                } else {
                    parents.push(last_top);
                }
            }
            let id_of = |i: usize| first_item + i;
            let siblings = |p: usize| -> Vec<usize> { (0..count).filter(|j| parents[*j] == p).collect() };
            let tops = siblings(usize::MAX);
            let mut items: Vec<Vec<u8>> = Vec::new();
            for (i, b) in self.bookmarks.iter().enumerate() {
                let sib = siblings(parents[i]);
                let pos = sib.iter().position(|x| *x == i).unwrap();
                let parent_id = if parents[i] == usize::MAX { root_id } else { id_of(parents[i]) };
                let mut d = format!(
                    "<< /Title {} /Parent {parent_id} 0 R /Dest [{} 0 R /XYZ 0 {:.2} 0]",
                    pdf_string(&b.title),
                    page_ids[b.page],
                    PAGE_H - b.y + 14.0
                );
                if pos > 0 {
                    d.push_str(&format!(" /Prev {} 0 R", id_of(sib[pos - 1])));
                }
                if pos + 1 < sib.len() {
                    d.push_str(&format!(" /Next {} 0 R", id_of(sib[pos + 1])));
                }
                let kids = siblings(i);
                if !kids.is_empty() {
                    d.push_str(&format!(" /First {} 0 R /Last {} 0 R /Count -{}", id_of(kids[0]), id_of(*kids.last().unwrap()), kids.len()));
                }
                d.push_str(" >>");
                items.push(d.into_bytes());
            }
            let root = format!(
                "<< /Type /Outlines /First {} 0 R /Last {} 0 R /Count {} >>",
                id_of(tops[0]),
                id_of(*tops.last().unwrap()),
                tops.len()
            );
            add(root.into_bytes(), &mut objects);
            for it in items {
                add(it, &mut objects);
            }
            outline_root = Some(root_id);
        }

        let outlines = outline_root.map(|id| format!(" /Outlines {id} 0 R /PageMode /UseOutlines")).unwrap_or_default();
        objects[0] = format!("<< /Type /Catalog /Pages 2 0 R{outlines} >>").into_bytes();
        let now = chrono::Local::now().format("D:%Y%m%d%H%M%S").to_string();
        let info_id = add(
            format!(
                "<< /Title {} /Author {} /Creator (Deployment Editor) /Producer (Deployment Editor {}) /CreationDate ({now}) >>",
                pdf_string(&self.title),
                pdf_string(&self.author),
                env!("CARGO_PKG_VERSION")
            )
            .into_bytes(),
            &mut objects,
        );

        let mut out: Vec<u8> = b"%PDF-1.4\n%\xE2\xE3\xCF\xD3\n".to_vec();
        let mut offsets = Vec::with_capacity(objects.len());
        for (i, obj) in objects.iter().enumerate() {
            offsets.push(out.len());
            out.extend_from_slice(format!("{} 0 obj\n", i + 1).as_bytes());
            out.extend_from_slice(obj);
            out.extend_from_slice(b"\nendobj\n");
        }
        let xref = out.len();
        out.extend_from_slice(format!("xref\n0 {}\n0000000000 65535 f \n", objects.len() + 1).as_bytes());
        for o in offsets {
            out.extend_from_slice(format!("{o:010} 00000 n \n").as_bytes());
        }
        out.extend_from_slice(format!("trailer\n<< /Size {} /Root 1 0 R /Info {info_id} 0 R >>\nstartxref\n{xref}\n%%EOF\n", objects.len() + 1).as_bytes());
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn widths_and_wrap() {
        assert!((text_width("Hello", Font::Regular, 10.0) - 22.78).abs() < 0.01); // H722 e556 l222 l222 o556
        assert_eq!(text_width("abc", Font::Mono, 10.0), 18.0);
        let lines = wrap("The quick brown fox jumps over the lazy dog", Font::Regular, 10.0, 80.0);
        assert!(lines.len() > 1 && lines.iter().all(|l| text_width(l, Font::Regular, 10.0) <= 80.0));
        let long = wrap("C:\\Program Files\\VeryLongFolderNameWithoutSpaces\\AndMore\\setup.exe", Font::Mono, 9.0, 100.0);
        assert!(long.iter().all(|l| text_width(l, Font::Mono, 9.0) <= 100.0));
        assert_eq!(winansi('ü'), 0xFC);
        assert_eq!(winansi('€'), 0x80);
        assert_eq!(pdf_string("a(b)\\ ä"), "(a\\(b\\)\\\\ \\344)");
    }

    #[test]
    fn writes_valid_structure() {
        let mut pdf = Pdf::new();
        pdf.title = "Test".into();
        pdf.heading("Section", 16.0, Rgb::hex(0x111111));
        for i in 0..120 {
            pdf.paragraph(&format!("Line {i} with some text that wraps maybe"), Font::Regular, 10.0, Rgb::hex(0), MARGIN_L, CONTENT_W);
        }
        pdf.code_block(&"Write-Host 'x'\n".repeat(80), MARGIN_L, CONTENT_W, 8.5, Rgb::hex(0xF3F4F6), Rgb::hex(0x111111));
        assert!(pdf.page_count() >= 3);
        let bytes = pdf.finish();
        let s = String::from_utf8_lossy(&bytes);
        assert!(s.starts_with("%PDF-1.4"));
        assert!(s.contains("/Outlines"));
        assert!(s.trim_end().ends_with("%%EOF"));
        // xref offsets point to "N 0 obj"
        let tail = &bytes[bytes.len() - 64..];
        let tail = String::from_utf8_lossy(tail);
        let xref_pos: usize = tail.rsplit("startxref\n").next().unwrap().lines().next().unwrap().parse().unwrap();
        assert!(bytes[xref_pos..].starts_with(b"xref"));
        let entry = &bytes[xref_pos..];
        let line3 = entry.split(|b| *b == b'\n').nth(3).unwrap();
        let off: usize = std::str::from_utf8(&line3[..10]).unwrap().parse().unwrap();
        assert!(bytes[off..].starts_with(b"1 0 obj"));
    }
}
