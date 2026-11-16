//! WinGet import: search packages with the `winget` CLI, read the installer
//! manifest from the public winget-pkgs repository and download the installer.
//!
//! The PureBasic version downloaded the complete winget-pkgs repository (hundreds
//! of MB) to search it. This port asks `winget search` instead and fetches only the
//! one manifest it needs from raw.githubusercontent.com.

use anyhow::{anyhow, bail, Result};
use std::path::Path;

use crate::system::{ps_quote, run_powershell_capture};

#[derive(Clone, Debug, PartialEq)]
pub struct Hit {
    pub name: String,
    pub id: String,
    pub version: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Installer {
    pub architecture: String,
    pub installer_type: String,
    pub url: String,
    pub silent: String,
    pub scope: String,
    pub product_code: String,
}

impl Installer {
    pub fn is_msi(&self) -> bool {
        matches!(self.installer_type.as_str(), "msi" | "wix") || self.url_extension() == "msi"
    }

    pub fn url_extension(&self) -> String {
        let path = self.url.split(['?', '#']).next().unwrap_or("");
        let file = path.rsplit('/').next().unwrap_or("");
        file.rsplit_once('.').map(|(_, e)| e.to_ascii_lowercase()).unwrap_or_default()
    }

    /// Extension used for `Files\Installer.<ext>`; `None` if not supported.
    pub fn local_extension(&self) -> Option<&'static str> {
        if self.is_msi() {
            return Some("msi");
        }
        match self.installer_type.as_str() {
            "exe" | "nullsoft" | "inno" | "burn" => Some("exe"),
            _ if self.url_extension() == "exe" => Some("exe"),
            _ => None,
        }
    }

    /// Silent switch from the manifest, or a well-known default for the installer technology.
    pub fn silent_switch(&self) -> String {
        if !self.silent.trim().is_empty() {
            return self.silent.trim().to_string();
        }
        match self.installer_type.as_str() {
            "nullsoft" => "/S".into(),
            "inno" => "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-".into(),
            "burn" => "/quiet /norestart".into(),
            _ => String::new(),
        }
    }

    pub fn label(&self) -> String {
        let scope = if self.scope.is_empty() { String::new() } else { format!(" · {}", self.scope) };
        format!("{} · {}{}", self.architecture, self.installer_type, scope)
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Manifest {
    pub id: String,
    pub version: String,
    pub installers: Vec<Installer>,
}

impl Manifest {
    /// Vendor and product name derived from the package identifier (e.g. `7zip.7zip`).
    pub fn vendor_and_name(&self) -> (String, String) {
        match self.id.split_once('.') {
            Some((vendor, rest)) => (vendor.to_string(), rest.replace('.', " ")),
            None => (self.id.clone(), self.id.clone()),
        }
    }

    /// Best installer for an architecture: machine scope and MSI/EXE preferred.
    pub fn best_for(&self, arch: &str) -> Option<usize> {
        let mut best: Option<(i32, usize)> = None;
        for (i, ins) in self.installers.iter().enumerate() {
            if ins.local_extension().is_none() {
                continue;
            }
            let mut score = 0;
            if ins.architecture.eq_ignore_ascii_case(arch) {
                score += 100;
            } else if ins.architecture == "neutral" {
                score += 50;
            } else {
                continue;
            }
            if ins.scope.is_empty() || ins.scope == "machine" {
                score += 10;
            }
            if !ins.silent_switch().is_empty() || ins.is_msi() {
                score += 5;
            }
            if best.map(|(s, _)| score > s).unwrap_or(true) {
                best = Some((score, i));
            }
        }
        best.map(|(_, i)| i)
    }
}

fn unquote(v: &str) -> String {
    let v = v.trim();
    let v = v
        .strip_prefix('\'')
        .and_then(|s| s.strip_suffix('\''))
        .or_else(|| v.strip_prefix('"').and_then(|s| s.strip_suffix('"')))
        .unwrap_or(v);
    v.replace("''", "'")
}

/// Minimal parser for the subset of the winget installer manifest we need.
pub fn parse_installer_manifest(yaml: &str) -> Manifest {
    let mut m = Manifest::default();
    let mut root = Installer::default();
    let mut current: Option<Installer> = None;
    let mut in_installers = false;
    let mut section = String::new(); // nested key such as InstallerSwitches
    let mut seq_indent: Option<usize> = None;

    for raw in yaml.lines() {
        let line = raw.trim_end();
        if line.trim().is_empty() || line.trim_start().starts_with('#') {
            continue;
        }
        let indent = line.len() - line.trim_start().len();
        let text = line.trim_start();

        if indent == 0 && !text.starts_with("- ") {
            section.clear();
            if let Some(c) = current.take() {
                m.installers.push(c);
            }
            in_installers = text.starts_with("Installers:");
            seq_indent = None;
            if let Some((k, v)) = text.split_once(':') {
                let v = unquote(v);
                match k {
                    "PackageIdentifier" => m.id = v,
                    "PackageVersion" => m.version = v,
                    "InstallerType" => root.installer_type = v,
                    "Scope" => root.scope = v,
                    "ProductCode" => root.product_code = v,
                    "InstallerSwitches" => section = "InstallerSwitches".into(),
                    _ => {}
                }
            }
            continue;
        }

        if !in_installers {
            if section == "InstallerSwitches" {
                if let Some(v) = text.strip_prefix("Silent:") {
                    root.silent = unquote(v);
                }
            }
            continue;
        }

        // Inside "Installers:" — items start with "- " at the sequence indent.
        // Nested lists (e.g. InstallModes) may use the same indent as item keys,
        // so the sequence indent is taken from the first item.
        let starts_item = text.starts_with("- ") && seq_indent.map(|s| s == indent).unwrap_or(true);
        let (body, is_key_line) = if starts_item {
            seq_indent.get_or_insert(indent);
            if let Some(c) = current.take() {
                m.installers.push(c);
            }
            current = Some(Installer::default());
            section.clear();
            (&text[2..], true)
        } else {
            (text, Some(indent) == seq_indent.map(|s| s + 2))
        };
        let Some(cur) = current.as_mut() else { continue };
        if is_key_line {
            section.clear();
            if let Some((k, v)) = body.split_once(':') {
                let v = unquote(v);
                match k.trim() {
                    "Architecture" => cur.architecture = v,
                    "InstallerType" => cur.installer_type = v,
                    "InstallerUrl" => cur.url = v.trim().to_string(),
                    "Scope" => cur.scope = v,
                    "ProductCode" => cur.product_code = v,
                    other => section = other.to_string(),
                }
            }
        } else if section == "InstallerSwitches" {
            if let Some(v) = body.trim_start_matches("- ").strip_prefix("Silent:") {
                cur.silent = unquote(v);
            }
        }
    }
    if let Some(c) = current.take() {
        m.installers.push(c);
    }
    // Root-level values are defaults for every installer.
    for ins in m.installers.iter_mut() {
        if ins.installer_type.is_empty() {
            ins.installer_type = root.installer_type.clone();
        }
        if ins.scope.is_empty() {
            ins.scope = root.scope.clone();
        }
        if ins.silent.is_empty() {
            ins.silent = root.silent.clone();
        }
        if ins.product_code.is_empty() {
            ins.product_code = root.product_code.clone();
        }
    }
    m
}

/// Parse the table printed by `winget search`.
pub fn parse_search_output(out: &str) -> Vec<Hit> {
    // Progress spinners are written with \r; keep only what follows the last \r.
    let lines: Vec<String> = out
        .lines()
        .map(|l| l.rsplit('\r').next().unwrap_or(l).to_string())
        .collect();
    let Some(sep) = lines.iter().position(|l| l.trim_start().starts_with("---") && l.trim().chars().all(|c| c == '-')) else {
        return Vec::new();
    };
    if sep == 0 {
        return Vec::new();
    }
    let header: Vec<char> = lines[sep - 1].chars().collect();
    // Column starts = positions where a word begins after 2+ spaces (or position 0).
    let mut starts = Vec::new();
    for i in 0..header.len() {
        if header[i] != ' ' && (i == 0 || (header[i - 1] == ' ' && (i < 2 || header[i - 2] == ' ' || starts.is_empty()))) {
            starts.push(i);
        }
    }
    if starts.len() < 3 {
        return Vec::new();
    }
    let col = |chars: &[char], idx: usize| -> String {
        let s = starts[idx].min(chars.len());
        let e = starts.get(idx + 1).copied().unwrap_or(chars.len()).min(chars.len());
        chars[s..e].iter().collect::<String>().trim().to_string()
    };
    let mut hits = Vec::new();
    for l in &lines[sep + 1..] {
        let chars: Vec<char> = l.chars().collect();
        if chars.iter().all(|c| c.is_whitespace()) {
            continue;
        }
        let name = col(&chars, 0);
        let id = col(&chars, 1);
        let version = col(&chars, 2);
        if id.is_empty() || id.contains(' ') {
            continue;
        }
        hits.push(Hit { name, id, version });
    }
    hits
}

pub fn manifest_url(id: &str, version: &str) -> String {
    let first = id.chars().next().map(|c| c.to_ascii_lowercase()).unwrap_or('_');
    let path = id.replace('.', "/");
    format!(
        "https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/{first}/{path}/{version}/{id}.installer.yaml"
    )
}

/// `winget search <query>` (blocking, Windows only).
pub fn search(query: &str) -> Result<Vec<Hit>> {
    let script = format!(
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; \
         if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {{ Write-Output 'WINGET-NOT-FOUND'; exit 2 }}; \
         winget search --query {} --source winget --accept-source-agreements --disable-interactivity",
        ps_quote(query)
    );
    let (code, out) = run_powershell_capture(&script)?;
    if out.contains("WINGET-NOT-FOUND") {
        bail!("The winget command was not found. Install 'App Installer' from the Microsoft Store, or enter the package identifier and version manually.");
    }
    let hits = parse_search_output(&out);
    if hits.is_empty() && code != 0 && !out.contains("No package found") {
        bail!("winget search failed:\n{}", out.trim());
    }
    Ok(hits)
}

fn download_to(url: &str, dest: &Path) -> Result<()> {
    let script = format!(
        "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; \
         Invoke-WebRequest -UseBasicParsing -Uri {} -OutFile {}",
        ps_quote(url),
        ps_quote(&dest.to_string_lossy())
    );
    let (code, out) = run_powershell_capture(&script)?;
    if code != 0 || !dest.is_file() {
        bail!("Download failed: {url}\n{}", out.trim());
    }
    Ok(())
}

/// Download and parse the installer manifest (blocking, Windows only).
pub fn fetch_manifest(id: &str, version: &str) -> Result<Manifest> {
    let tmp = std::env::temp_dir().join(format!("de-winget-{}-{}.yaml", id, version));
    download_to(&manifest_url(id, version), &tmp)
        .map_err(|e| anyhow!("Could not load the manifest for {id} {version}. Check identifier and version.\n{e}"))?;
    let text = std::fs::read_to_string(&tmp)?;
    let _ = std::fs::remove_file(&tmp);
    let m = parse_installer_manifest(&text);
    if m.installers.is_empty() {
        bail!("The manifest of {id} {version} contains no installers.");
    }
    Ok(m)
}

pub fn download_installer(url: &str, dest: &Path) -> Result<()> {
    download_to(url, dest)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_7zip_manifest() {
        let yaml = include_str!("../tests/data/7zip.installer.yaml");
        let m = parse_installer_manifest(yaml);
        assert_eq!(m.id, "7zip.7zip");
        assert_eq!(m.version, "25.01");
        assert_eq!(m.installers.len(), 6);
        let x64 = &m.installers[1];
        assert_eq!(x64.architecture, "x64");
        assert_eq!(x64.installer_type, "exe");
        assert_eq!(x64.url, "https://7-zip.org/a/7z2501-x64.exe");
        assert_eq!(x64.silent, "/S");
        assert_eq!(x64.scope, "machine", "root scope inherited");
        let msi = &m.installers[5];
        assert!(msi.is_msi());
        assert_eq!(msi.product_code, "{23170F69-40C1-2702-2501-000001000000}");
        let best = m.best_for("x64").unwrap();
        assert_eq!(m.installers[best].architecture, "x64");
        assert_eq!(m.vendor_and_name(), ("7zip".to_string(), "7zip".to_string()));
    }

    #[test]
    fn root_level_type_and_switches() {
        let yaml = "PackageIdentifier: Vendor.App\nPackageVersion: 1.2.3\nInstallerType: inno\nInstallerSwitches:\n  Silent: /SILENT\nInstallers:\n- Architecture: x64\n  InstallerUrl: https://x/y/setup.exe?dl=1\nManifestType: installer\n";
        let m = parse_installer_manifest(yaml);
        assert_eq!(m.installers.len(), 1);
        let i = &m.installers[0];
        assert_eq!(i.installer_type, "inno");
        assert_eq!(i.silent_switch(), "/SILENT");
        assert_eq!(i.url_extension(), "exe");
        assert_eq!(i.local_extension(), Some("exe"));
    }

    #[test]
    fn parses_search_table() {
        let out = "   - \r   \\ \r\nName                 Id                       Version   Match        Source\r\n-----------------------------------------------------------------------------\r\n7-Zip                7zip.7zip                25.01     Moniker: 7zip winget\r\n7-Zip ZS             mcmilk.7zip-zstd         24.09.0.0 Tag: 7zip    winget\r\nNanaZip              M2Team.NanaZip           5.0.1263.0 Tag: 7-zip  winget\r\n";
        let hits = parse_search_output(out);
        assert_eq!(hits.len(), 3);
        assert_eq!(hits[0], Hit { name: "7-Zip".into(), id: "7zip.7zip".into(), version: "25.01".into() });
        assert_eq!(hits[1].id, "mcmilk.7zip-zstd");
        assert_eq!(hits[1].name, "7-Zip ZS");
    }

    #[test]
    fn manifest_urls() {
        assert_eq!(
            manifest_url("Notepad++.Notepad++", "8.8.1"),
            "https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/n/Notepad++/Notepad++/8.8.1/Notepad++.Notepad++.installer.yaml"
        );
    }
}
