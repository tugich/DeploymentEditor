//! Plugins (`Plugins/Enabled-*/Plugin.ini`) and PowerShell snippets (`Snippets/*.ps1`).

use std::fs;
use std::path::{Path, PathBuf};

#[derive(Clone, Debug, Default)]
pub struct Plugin {
    pub id: String,
    pub name: String,
    pub version: String,
    pub date: String,
    pub description: String,
    pub author: String,
    pub website: String,
    pub folder: PathBuf,
    pub file: String,
    pub parameter: String,
    pub unload_project: bool,
}

impl Plugin {
    pub fn script_path(&self) -> PathBuf {
        self.folder.join(&self.file)
    }
}

/// Minimal INI reader compatible with PureBasic's preference files.
fn parse_ini(text: &str) -> Vec<(String, String, String)> {
    let mut section = String::new();
    let mut out = Vec::new();
    for raw in text.lines() {
        let line = raw.trim().trim_start_matches('\u{feff}');
        if line.is_empty() || line.starts_with(';') || line.starts_with('#') {
            continue;
        }
        if line.starts_with('[') && line.ends_with(']') {
            section = line[1..line.len() - 1].trim().to_string();
            continue;
        }
        if let Some((k, v)) = line.split_once('=') {
            out.push((section.clone(), k.trim().to_string(), v.trim().to_string()));
        }
    }
    out
}

pub fn load_plugins(dir: &Path) -> Vec<Plugin> {
    let mut plugins = Vec::new();
    let Ok(entries) = fs::read_dir(dir) else { return plugins };
    for entry in entries.flatten() {
        let path = entry.path();
        if !path.is_dir() {
            continue;
        }
        let folder_name = entry.file_name().to_string_lossy().to_string();
        let Some(id) = folder_name.strip_prefix("Enabled-") else { continue };
        let ini = fs::read_to_string(path.join("Plugin.ini")).unwrap_or_default();
        let mut p = Plugin { id: id.to_string(), folder: path.clone(), ..Default::default() };
        for (section, key, value) in parse_ini(&ini) {
            match (section.as_str(), key.as_str()) {
                ("Details", "Name") => p.name = value,
                ("Details", "Version") => p.version = value,
                ("Details", "Date") => p.date = value,
                ("Details", "Description") => p.description = value,
                ("Details", "Author") => p.author = value,
                ("Details", "Website") => p.website = value,
                ("Script", "File") => p.file = value,
                ("Script", "Parameter") => p.parameter = value,
                ("Script", "UnloadProject") => p.unload_project = value.eq_ignore_ascii_case("yes"),
                _ => {}
            }
        }
        if p.name.is_empty() {
            p.name = p.id.replace('_', " ");
        }
        plugins.push(p);
    }
    plugins.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
    plugins
}

#[derive(Clone, Debug)]
pub struct Snippet {
    pub file_name: String,
    pub path: PathBuf,
}

pub fn load_snippets(dir: &Path) -> Vec<Snippet> {
    let mut list: Vec<Snippet> = fs::read_dir(dir)
        .into_iter()
        .flatten()
        .flatten()
        .filter(|e| {
            e.path().is_file()
                && e.path().extension().map(|x| x.eq_ignore_ascii_case("ps1")).unwrap_or(false)
        })
        .map(|e| Snippet { file_name: e.file_name().to_string_lossy().to_string(), path: e.path() })
        .collect();
    list.sort_by(|a, b| a.file_name.to_lowercase().cmp(&b.file_name.to_lowercase()));
    list
}

/// Read a text file the way PureBasic's `ReadStringFormat()` did: honour a BOM
/// (UTF-8 / UTF-16 LE / BE), otherwise assume UTF-8. Line endings are CRLF.
pub fn read_text_file(path: &Path) -> std::io::Result<String> {
    let bytes = fs::read(path)?;
    let text = if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        String::from_utf8_lossy(&bytes[3..]).to_string()
    } else if bytes.starts_with(&[0xFF, 0xFE]) {
        let u: Vec<u16> = bytes[2..].chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
        String::from_utf16_lossy(&u)
    } else if bytes.starts_with(&[0xFE, 0xFF]) {
        let u: Vec<u16> = bytes[2..].chunks_exact(2).map(|c| u16::from_be_bytes([c[0], c[1]])).collect();
        String::from_utf16_lossy(&u)
    } else {
        String::from_utf8_lossy(&bytes).to_string()
    };
    Ok(to_crlf(&text))
}

pub fn to_crlf(s: &str) -> String {
    s.replace("\r\n", "\n").replace('\n', "\r\n")
}
