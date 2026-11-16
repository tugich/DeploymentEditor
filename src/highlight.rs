//! A small PowerShell syntax highlighter for the script preview and custom-script editor.

use eframe::egui::{text::LayoutJob, Color32, FontId, TextFormat};

#[derive(Clone, Copy)]
struct Colors {
    text: Color32,
    comment: Color32,
    string: Color32,
    variable: Color32,
    command: Color32,
    parameter: Color32,
    keyword: Color32,
    number: Color32,
    operator: Color32,
}

fn colors(dark: bool) -> Colors {
    if dark {
        Colors {
            text: Color32::from_rgb(0xD4, 0xD4, 0xD4),
            comment: Color32::from_rgb(0x6A, 0x99, 0x55),
            string: Color32::from_rgb(0xCE, 0x91, 0x78),
            variable: Color32::from_rgb(0x9C, 0xDC, 0xFE),
            command: Color32::from_rgb(0xDC, 0xDC, 0xAA),
            parameter: Color32::from_rgb(0x8F, 0xA8, 0xC8),
            keyword: Color32::from_rgb(0xC5, 0x86, 0xC0),
            number: Color32::from_rgb(0xB5, 0xCE, 0xA8),
            operator: Color32::from_rgb(0xA0, 0xA0, 0xA0),
        }
    } else {
        Colors {
            text: Color32::from_rgb(0x1F, 0x23, 0x28),
            comment: Color32::from_rgb(0x00, 0x80, 0x00),
            string: Color32::from_rgb(0xA3, 0x15, 0x15),
            variable: Color32::from_rgb(0x00, 0x10, 0x80),
            command: Color32::from_rgb(0x79, 0x5E, 0x26),
            parameter: Color32::from_rgb(0x00, 0x5C, 0xC5),
            keyword: Color32::from_rgb(0xAF, 0x00, 0xDB),
            number: Color32::from_rgb(0x09, 0x86, 0x58),
            operator: Color32::from_rgb(0x55, 0x55, 0x55),
        }
    }
}

const KEYWORDS: &[&str] = &[
    "begin", "break", "catch", "class", "continue", "data", "do", "dynamicparam", "else", "elseif", "end",
    "exit", "filter", "finally", "for", "foreach", "function", "if", "in", "param", "process", "return",
    "switch", "throw", "trap", "try", "until", "using", "while",
];

pub fn highlight_powershell(text: &str, font: FontId, dark: bool) -> LayoutJob {
    let c = colors(dark);
    let mut job = LayoutJob::default();
    let fmt = |color: Color32| TextFormat { font_id: font.clone(), color, ..Default::default() };

    let bytes = text.as_bytes();
    let mut i = 0;
    let n = bytes.len();
    let push = |job: &mut LayoutJob, s: &str, color: Color32| job.append(s, 0.0, fmt(color));

    while i < n {
        let ch = text[i..].chars().next().unwrap();
        let len = ch.len_utf8();

        // Block comment <# ... #>
        if text[i..].starts_with("<#") {
            let end = text[i + 2..].find("#>").map(|e| i + 2 + e + 2).unwrap_or(n);
            push(&mut job, &text[i..end], c.comment);
            i = end;
            continue;
        }
        // Line comment
        if ch == '#' {
            let end = text[i..].find('\n').map(|e| i + e).unwrap_or(n);
            push(&mut job, &text[i..end], c.comment);
            i = end;
            continue;
        }
        // Strings
        if ch == '\'' || ch == '"' {
            let mut j = i + 1;
            while j < n {
                let cj = bytes[j] as char;
                if ch == '"' && cj == '`' {
                    j += 2;
                    continue;
                }
                if cj == ch {
                    if j + 1 < n && bytes[j + 1] as char == ch {
                        j += 2; // escaped quote
                        continue;
                    }
                    j += 1;
                    break;
                }
                j += 1;
            }
            let j = j.min(n);
            let j = (j..=n).find(|k| text.is_char_boundary(*k)).unwrap_or(n);
            push(&mut job, &text[i..j], c.string);
            i = j;
            continue;
        }
        // Variables
        if ch == '$' {
            let mut j = i + 1;
            if j < n && bytes[j] == b'{' {
                j = text[j..].find('}').map(|e| j + e + 1).unwrap_or(n);
            } else {
                while j < n {
                    let cj = text[j..].chars().next().unwrap();
                    if cj.is_alphanumeric() || cj == '_' || cj == ':' || cj == '?' {
                        j += cj.len_utf8();
                    } else {
                        break;
                    }
                }
            }
            push(&mut job, &text[i..j], c.variable);
            i = j;
            continue;
        }
        // Parameters / operators starting with '-'
        if ch == '-' && i + 1 < n && (bytes[i + 1] as char).is_ascii_alphabetic() {
            let prev_ok = i == 0 || {
                let p = bytes[i - 1] as char;
                p.is_whitespace() || p == '(' || p == '{'
            };
            if prev_ok {
                let mut j = i + 1;
                while j < n && ((bytes[j] as char).is_ascii_alphanumeric() || bytes[j] == b'_') {
                    j += 1;
                }
                push(&mut job, &text[i..j], c.parameter);
                i = j;
                continue;
            }
        }
        // Numbers
        if ch.is_ascii_digit() {
            let mut j = i;
            while j < n && ((bytes[j] as char).is_ascii_alphanumeric() || bytes[j] == b'.') {
                j += 1;
            }
            push(&mut job, &text[i..j], c.number);
            i = j;
            continue;
        }
        // Words: commands (Verb-Noun), keywords
        if ch.is_alphabetic() || ch == '_' {
            let mut j = i;
            while j < n {
                let cj = text[j..].chars().next().unwrap();
                if cj.is_alphanumeric() || cj == '_' || cj == '-' || cj == '.' {
                    j += cj.len_utf8();
                } else {
                    break;
                }
            }
            let word = &text[i..j];
            let color = if KEYWORDS.contains(&word.to_ascii_lowercase().as_str()) {
                c.keyword
            } else if word.contains('-')
                && word.chars().next().map(|x| x.is_uppercase()).unwrap_or(false)
                && !word.ends_with('-')
            {
                c.command
            } else {
                c.text
            };
            push(&mut job, word, color);
            i = j;
            continue;
        }
        if "=+*/%!<>|&;,(){}[]@".contains(ch) {
            push(&mut job, &text[i..i + len], c.operator);
            i += len;
            continue;
        }
        // whitespace / anything else
        let mut j = i + len;
        while j < n {
            let cj = text[j..].chars().next().unwrap();
            if cj.is_whitespace() {
                j += cj.len_utf8();
            } else {
                break;
            }
        }
        push(&mut job, &text[i..j], c.text);
        i = j;
    }
    job
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn covers_whole_text() {
        let src = "<# hi #>\n$a = Start-ADTProcess -FilePath \"x`\"y\" 'it''s' # c\nif ($true) { 42 } ü";
        let job = highlight_powershell(src, FontId::monospace(12.0), false);
        assert_eq!(job.text, src);
    }
}
