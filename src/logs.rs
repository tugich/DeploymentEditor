//! Parser for CMTrace-style (`<![LOG[...]LOG]!>`) and plain-text log files,
//! used by the built-in Software Log Reader.

use std::path::{Path, PathBuf};

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum Severity {
    Info,
    Warning,
    Error,
}

#[derive(Clone, Debug)]
pub struct LogEntry {
    pub file: usize,
    pub time: String,
    pub date: String,
    pub component: String,
    pub severity: Severity,
    pub message: String,
}

#[derive(Default)]
pub struct LogSet {
    pub files: Vec<PathBuf>,
    pub entries: Vec<LogEntry>,
}

fn attr(meta: &str, key: &str) -> String {
    let pat = format!("{key}=\"");
    meta.find(&pat)
        .and_then(|s| {
            let rest = &meta[s + pat.len()..];
            rest.find('"').map(|e| rest[..e].to_string())
        })
        .unwrap_or_default()
}

fn guess_severity(line: &str) -> Severity {
    let l = line.to_ascii_lowercase();
    if l.contains("error") || l.contains("failed") || l.contains("exception") {
        Severity::Error
    } else if l.contains("warn") {
        Severity::Warning
    } else {
        Severity::Info
    }
}

pub fn parse_log(text: &str, file: usize) -> Vec<LogEntry> {
    let mut out = Vec::new();
    if text.contains("<![LOG[") {
        let mut rest = text;
        while let Some(start) = rest.find("<![LOG[") {
            let after = &rest[start + 7..];
            let Some(msg_end) = after.find("]LOG]!>") else { break };
            let message = after[..msg_end].trim_end().to_string();
            let meta_start = &after[msg_end + 7..];
            let meta_end = meta_start.find('>').unwrap_or(meta_start.len());
            let meta = &meta_start[..meta_end];
            let severity = match attr(meta, "type").as_str() {
                "2" => Severity::Warning,
                "3" => Severity::Error,
                _ => Severity::Info,
            };
            let mut time = attr(meta, "time");
            if let Some(cut) = time.find(['+', '-']) {
                time.truncate(cut);
            }
            if let Some(dot) = time.find('.') {
                time.truncate((dot + 4).min(time.len()));
            }
            out.push(LogEntry {
                file,
                time,
                date: attr(meta, "date"),
                component: attr(meta, "component"),
                severity,
                message,
            });
            rest = &meta_start[meta_end.min(meta_start.len())..];
        }
    } else {
        for line in text.lines() {
            if line.trim().is_empty() {
                continue;
            }
            // Legacy PSADT format: [date time] [component] [Severity] :: message
            let (meta, message) = match line.split_once(" :: ") {
                Some((m, msg)) => (m, msg.to_string()),
                None => ("", line.to_string()),
            };
            let parts: Vec<&str> = meta
                .split(']')
                .map(|s| s.trim().trim_start_matches('['))
                .filter(|s| !s.is_empty())
                .collect();
            let severity = parts
                .iter()
                .find_map(|p| match p.to_ascii_lowercase().as_str() {
                    "error" => Some(Severity::Error),
                    "warning" => Some(Severity::Warning),
                    "info" | "success" => Some(Severity::Info),
                    _ => None,
                })
                .unwrap_or_else(|| guess_severity(line));
            let (date, time) = parts
                .first()
                .and_then(|p| p.split_once(' '))
                .map(|(d, t)| (d.to_string(), t.to_string()))
                .unwrap_or_default();
            out.push(LogEntry {
                file,
                time,
                date,
                component: parts.get(1).map(|s| s.to_string()).unwrap_or_default(),
                severity,
                message,
            });
        }
    }
    out
}

pub fn read_log_file(path: &Path) -> std::io::Result<String> {
    // logs may be UTF-16 (CMTrace), UTF-8 with BOM, or ANSI
    let text = crate::plugins::read_text_file(path)?;
    Ok(text.replace("\r\n", "\n"))
}

impl LogSet {
    pub fn add_file(&mut self, path: PathBuf) -> std::io::Result<usize> {
        let text = read_log_file(&path)?;
        let idx = self.files.len();
        let mut entries = parse_log(&text, idx);
        let n = entries.len();
        self.entries.append(&mut entries);
        self.files.push(path);
        Ok(n)
    }

    pub fn reload(&mut self) {
        let files = std::mem::take(&mut self.files);
        self.entries.clear();
        for f in files {
            let _ = self.add_file(f);
        }
    }
}

/// Default PSADT log folder.
pub fn default_log_folder() -> PathBuf {
    let windir = std::env::var("WINDIR").unwrap_or_else(|_| "C:\\Windows".into());
    PathBuf::from(windir).join("Logs").join("Software")
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn cmtrace() {
        let t = r#"<![LOG[Hello world]LOG]!><time="10:11:12.3456789+60" date="11-07-2026" component="Install" context="" type="3" thread="1" file="x">
<![LOG[Second
line]LOG]!><time="10:11:13.000+60" date="11-07-2026" component="Pre" context="" type="1" thread="1" file="x">"#;
        let e = parse_log(t, 0);
        assert_eq!(e.len(), 2);
        assert_eq!(e[0].severity, Severity::Error);
        assert_eq!(e[0].time, "10:11:12.345");
        assert_eq!(e[1].message, "Second\nline");
        assert_eq!(e[1].component, "Pre");
    }
    #[test]
    fn legacy() {
        let e = parse_log("[2026-11-07 10:00:00.123] [Installation] [Warning] :: Something odd", 0);
        assert_eq!(e[0].severity, Severity::Warning);
        assert_eq!(e[0].component, "Installation");
        assert_eq!(e[0].message, "Something odd");
    }
}
