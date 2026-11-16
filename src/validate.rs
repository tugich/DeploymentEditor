//! Checks that run before generating, testing or uploading a package.
//!
//! Errors block Generate / Run / Upload (the user can still continue anyway);
//! warnings and hints are only shown.

use std::path::Path;

use crate::catalog::Catalog;
use crate::detection::DetectionRule;
use crate::project::{keys, Action, DeploymentType, Project};

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum Severity {
    Error,
    Warning,
    Hint,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Issue {
    pub severity: Severity,
    pub dt: Option<DeploymentType>,
    pub action_id: Option<i64>,
    pub step: Option<usize>,
    pub action_name: String,
    pub message: String,
}

impl Issue {
    fn project(severity: Severity, message: impl Into<String>) -> Self {
        Self { severity, dt: None, action_id: None, step: None, action_name: String::new(), message: message.into() }
    }
    fn step(severity: Severity, dt: DeploymentType, idx: usize, a: &Action, message: impl Into<String>) -> Self {
        Self {
            severity,
            dt: Some(dt),
            action_id: Some(a.id),
            step: Some(idx + 1),
            action_name: a.name.clone(),
            message: message.into(),
        }
    }
    /// Where the issue is, e.g. "Installation · step 3 · Install 7-Zip".
    pub fn location(&self) -> String {
        match (self.dt, self.step) {
            (Some(dt), Some(n)) => format!("{} · step {n} · {}", dt.as_db(), self.action_name),
            (Some(dt), None) => dt.as_db().to_string(),
            _ => "Project".into(),
        }
    }
}

/// Parameters that a step really needs. Each inner list is "one of these".
/// The `Required` flag in the PSADT database means "mandatory in at least one
/// parameter set" (it even includes switches), so it is only used as a soft hint.
fn key_parameters(command: &str) -> &'static [&'static [&'static str]] {
    match command {
        "Start-ADTProcess" | "Start-ADTMspProcess" | "Start-ADTProcessAsUser" => &[&["FilePath"]],
        "Start-ADTMsiProcess" => &[&["FilePath", "ProductCode", "InstalledApplication"]],
        "Uninstall-ADTApplication" => &[&["Name", "ProductCode", "FilterScript", "InstalledApplication"]],
        "Copy-ADTFile" => &[&["Path"], &["Destination"]],
        "Copy-ADTFileToUserProfiles" => &[&["Path"]],
        "New-ADTFolder" => &[&["LiteralPath"]],
        "Remove-ADTFolder" => &[&["Path", "LiteralPath"]],
        "New-ADTShortcut" => &[&["LiteralPath"], &["TargetPath"]],
        "Remove-ADTFile" => &[&["Path", "LiteralPath"]],
        "Remove-ADTRegistryKey" => &[&["Path", "LiteralPath"]],
        "Set-ADTRegistryKey" => &[&["LiteralPath"]],
        "Set-ADTIniValue" => &[&["FilePath"], &["Section"], &["Key"]],
        "Show-ADTInstallationPrompt" => &[&["Message"]],
        "Show-ADTDialogBox" => &[&["Text"]],
        "Block-ADTAppExecution" => &[&["Processes"]],
        _ => &[],
    }
}

/// Programs that are found on the target machine's PATH.
const SYSTEM_PROGRAMS: &[&str] = &[
    "msiexec", "cmd", "powershell", "pwsh", "reg", "regsvr32", "rundll32", "sc", "net", "netsh", "dism", "wusa",
    "icacls", "takeown", "taskkill", "robocopy", "xcopy", "schtasks", "bcdedit", "certutil", "wmic", "expand",
];

/// Finds unbalanced brackets / quotes in PowerShell code. Strings, comments and
/// here-strings are skipped. Returns a human readable problem.
pub fn bracket_problem(code: &str) -> Option<String> {
    let chars: Vec<char> = code.chars().collect();
    let mut stack: Vec<(char, usize)> = Vec::new();
    let mut line = 1usize;
    let mut i = 0;
    let n = chars.len();
    while i < n {
        let c = chars[i];
        if c == '\n' {
            line += 1;
            i += 1;
            continue;
        }
        // here-strings @" ... "@ and @' ... '@ (closing at line start)
        if c == '@' && i + 1 < n && (chars[i + 1] == '"' || chars[i + 1] == '\'') {
            let q = chars[i + 1];
            let start = line;
            let mut j = i + 2;
            let mut found = false;
            while j + 1 < n {
                if chars[j] == '\n' {
                    line += 1;
                    if chars[j + 1] == q && j + 2 < n && chars[j + 2] == '@' {
                        j += 3;
                        found = true;
                        break;
                    }
                }
                j += 1;
            }
            if !found {
                return Some(format!("Here-string starting on line {start} is not closed"));
            }
            i = j;
            continue;
        }
        if c == '<' && i + 1 < n && chars[i + 1] == '#' {
            let start = line;
            let mut j = i + 2;
            while j + 1 < n && !(chars[j] == '#' && chars[j + 1] == '>') {
                if chars[j] == '\n' {
                    line += 1;
                }
                j += 1;
            }
            if j + 1 >= n {
                return Some(format!("Block comment starting on line {start} is not closed"));
            }
            i = j + 2;
            continue;
        }
        if c == '#' {
            while i < n && chars[i] != '\n' {
                i += 1;
            }
            continue;
        }
        if c == '`' {
            i += 2; // escaped character
            continue;
        }
        if c == '\'' || c == '"' {
            let start = line;
            let mut j = i + 1;
            let mut closed = false;
            while j < n {
                let d = chars[j];
                if d == '\n' {
                    line += 1;
                }
                if c == '"' && d == '`' {
                    j += 2;
                    continue;
                }
                if d == c {
                    if j + 1 < n && chars[j + 1] == c {
                        j += 2; // doubled quote = escaped
                        continue;
                    }
                    closed = true;
                    j += 1;
                    break;
                }
                j += 1;
            }
            if !closed {
                let kind = if c == '"' { "Double" } else { "Single" };
                return Some(format!("{kind}-quoted string starting on line {start} is not closed"));
            }
            i = j;
            continue;
        }
        match c {
            '(' | '{' | '[' => stack.push((c, line)),
            ')' | '}' | ']' => {
                let want = match c {
                    ')' => '(',
                    '}' => '{',
                    _ => '[',
                };
                match stack.pop() {
                    Some((open, _)) if open == want => {}
                    Some((open, l)) => {
                        return Some(format!("'{c}' on line {line} does not match '{open}' from line {l}"));
                    }
                    None => return Some(format!("'{c}' on line {line} has no opening bracket")),
                }
            }
            _ => {}
        }
        i += 1;
    }
    stack.pop().map(|(open, l)| format!("'{open}' opened on line {l} is never closed"))
}

fn is_valid_variable(v: &str) -> bool {
    let v = v.trim();
    let Some(rest) = v.strip_prefix('$') else { return false };
    let rest = rest.strip_prefix("script:").or_else(|| rest.strip_prefix("global:")).unwrap_or(rest);
    !rest.is_empty() && rest.chars().all(|c| c.is_alphanumeric() || c == '_')
}

fn has_unescaped_double_quote(v: &str) -> bool {
    let chars: Vec<char> = v.chars().collect();
    chars.iter().enumerate().any(|(i, c)| *c == '"' && (i == 0 || chars[i - 1] != '`'))
}

/// Path relative to the package (resolved by PSADT against the Files folder)?
fn is_relative_file(v: &str) -> bool {
    let v = v.trim();
    !v.is_empty()
        && !v.contains('$')
        && !v.contains('%')
        && !v.contains(':')
        && !v.starts_with('\\')
        && !v.starts_with('/')
        && !v.starts_with('(')
}

fn check_action(issues: &mut Vec<Issue>, dt: DeploymentType, idx: usize, a: &Action, catalog: &Catalog, project_folder: &Path) {
    let mut push = |sev: Severity, msg: String| issues.push(Issue::step(sev, dt, idx, a, msg));

    if a.is_custom_script() {
        let script = a.value("Script");
        if script.trim().is_empty() {
            push(Severity::Warning, "The custom script is empty.".into());
        } else if let Some(p) = bracket_problem(script) {
            push(Severity::Error, format!("Script: {p}."));
        }
    } else if catalog.command(&a.command).is_none() {
        push(Severity::Warning, format!("{} is not in the PSADT command database.", a.command));
    } else {
        let set = |name: &str| {
            let v = a.value(name).trim().to_string();
            let is_switch = catalog.parameter(&a.command, name).map(|p| p.is_checkbox()).unwrap_or(false);
            if is_switch { v == "1" } else { !v.is_empty() }
        };
        let curated = key_parameters(&a.command);
        for group in curated {
            if !group.iter().any(|p| set(p)) {
                let msg = if group.len() == 1 {
                    format!("Parameter -{} is required.", group[0])
                } else {
                    format!("One of -{} is required.", group.join(", -"))
                };
                push(Severity::Error, msg);
            }
        }
        if curated.is_empty() {
            let required: Vec<&str> = catalog
                .parameters(&a.command)
                .iter()
                .filter(|p| p.required && !p.is_checkbox())
                .map(|p| p.name.as_str())
                .collect();
            if !required.is_empty() && !required.iter().any(|p| set(p)) {
                push(Severity::Warning, format!("Probably needs one of: -{}.", required.join(", -")));
            }
        }

        for (name, value) in &a.values {
            let v = value.trim();
            if v.is_empty() {
                continue;
            }
            let Some(info) = catalog.parameter(&a.command, name) else { continue };
            if info.is_checkbox() {
                continue;
            }
            if info.ptype.contains("String") && !info.ptype.contains("[]") && has_unescaped_double_quote(v) {
                push(
                    Severity::Error,
                    format!("-{name} contains a double quote. Escape it as `\" or use single quotes inside the value."),
                );
            }
            if matches!(info.ptype.as_str(), "Int32" | "UInt32" | "Int64" | "Double")
                && !v.starts_with('$')
                && !v.starts_with('(')
                && v.parse::<f64>().is_err()
            {
                push(Severity::Error, format!("-{name} expects a number, but is \"{v}\"."));
            }
            if !info.ptype.contains("String") && !v.starts_with('"') && !v.starts_with('\'') {
                if let Some(p) = bracket_problem(v) {
                    push(Severity::Error, format!("-{name}: {p}."));
                }
            }
        }

        // Installer files referenced relative to the package
        if matches!(a.command.as_str(), "Start-ADTProcess" | "Start-ADTMsiProcess" | "Start-ADTMspProcess") {
            let file = a.value("FilePath").trim().trim_matches('"').trim_matches('\'');
            if is_relative_file(file) {
                let in_files = project_folder.join("Files").join(file);
                let in_root = project_folder.join(file);
                if !in_files.exists() && !in_root.exists() {
                    let stem = Path::new(file)
                        .file_stem()
                        .map(|s| s.to_string_lossy().to_ascii_lowercase())
                        .unwrap_or_default();
                    let is_system = a.command == "Start-ADTProcess" && !file.contains('\\') && SYSTEM_PROGRAMS.contains(&stem.as_str());
                    if !is_system {
                        let sev = if a.command == "Start-ADTProcess" { Severity::Warning } else { Severity::Error };
                        push(sev, format!("\"{file}\" was not found in the Files folder of the project."));
                    }
                }
            }
        }
    }

    if !a.conditions.trim().is_empty() {
        if let Some(p) = bracket_problem(&a.conditions) {
            push(Severity::Error, format!("Condition: {p}."));
        } else {
            for leaf in crate::condition::parse(&a.conditions).leaves() {
                for p in leaf.problems() {
                    push(Severity::Warning, format!("Condition \"{}\": {p}", leaf.describe()));
                }
            }
        }
    }
    if !a.variable_name.trim().is_empty() && !is_valid_variable(&a.variable_name) {
        push(Severity::Error, format!("\"{}\" is not a valid variable name (e.g. $result).", a.variable_name.trim()));
    }
}

/// Run all checks for a project.
pub fn validate(project: &Project, catalog: &Catalog) -> Vec<Issue> {
    let mut issues = Vec::new();

    for (key, label) in [(keys::APP_NAME, "Application name"), (keys::APP_VERSION, "Version"), (keys::APP_VENDOR, "Vendor")] {
        if project.setting(key).trim().is_empty() {
            issues.push(Issue::project(Severity::Warning, format!("{label} is empty (Project settings).")));
        }
    }

    for dt in DeploymentType::ALL {
        let actions = project.actions(dt).unwrap_or_default();
        let enabled = actions.iter().filter(|a| !a.disabled).count();
        match dt {
            DeploymentType::Installation if enabled == 0 => issues.push(Issue {
                dt: Some(dt),
                ..Issue::project(Severity::Warning, "The installation sequence has no active steps.")
            }),
            DeploymentType::Uninstall if enabled == 0 => issues.push(Issue {
                dt: Some(dt),
                ..Issue::project(Severity::Warning, "The uninstall sequence is empty — Intune needs a working uninstall.")
            }),
            _ => {}
        }
        for (idx, a) in actions.iter().enumerate() {
            if !a.disabled {
                check_action(&mut issues, dt, idx, a, catalog, &project.folder);
            }
        }
    }

    match DetectionRule::load(project) {
        None => issues.push(Issue::project(Severity::Hint, "No detection rule yet — needed for Upload to Intune (Intune tab).")),
        Some(rule) => {
            for p in rule.problems() {
                issues.push(Issue::project(Severity::Warning, format!("Detection rule: {p}")));
            }
        }
    }

    issues.sort_by_key(|i| (i.severity, i.dt.map(|d| d as u8).unwrap_or(0), i.step.unwrap_or(0)));
    issues
}

pub fn count(issues: &[Issue], sev: Severity) -> usize {
    issues.iter().filter(|i| i.severity == sev).count()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn brackets() {
        assert_eq!(bracket_problem("if ($a -eq 1) { Write-Host 'x)' }"), None);
        assert_eq!(bracket_problem("$x = \"a `\" b\"; # comment ("), None);
        assert_eq!(bracket_problem("$h = @\"\nline (\n\"@\n$y = 1"), None);
        assert_eq!(bracket_problem("<# ( #> $a"), None);
        assert_eq!(bracket_problem("'it''s'"), None);
        assert!(bracket_problem("foreach ($i in 1..3) {\n Write-Host $i\n").unwrap().contains("line 1"));
        assert!(bracket_problem("Get-Item 'abc").unwrap().contains("not closed"));
        assert!(bracket_problem("($a]").unwrap().contains("does not match"));
        assert!(bracket_problem(")").is_some());
    }

    #[test]
    fn variables_and_quotes() {
        assert!(is_valid_variable("$result"));
        assert!(is_valid_variable("$script:Result_1"));
        assert!(!is_valid_variable("result"));
        assert!(!is_valid_variable("$res ult"));
        assert!(has_unescaped_double_quote("say \"hi\""));
        assert!(!has_unescaped_double_quote("say `\"hi`\""));
        assert!(is_relative_file("setup.exe"));
        assert!(is_relative_file("Sub\\setup.msi"));
        assert!(!is_relative_file("$adtSession.DirFiles\\setup.exe"));
        assert!(!is_relative_file("C:\\x.exe"));
    }

    #[test]
    fn demo_project_has_no_errors() {
        let Ok(root) = std::env::var("DE_REPO") else { return };
        let paths = crate::paths::AppPaths::with_root(&root);
        let catalog = Catalog::load(&paths.psadt_database()).unwrap();
        let dir = std::env::temp_dir().join("de-validate-demo");
        let _ = std::fs::remove_dir_all(&dir);
        crate::project::copy_dir_all(&paths.demo_project(), &dir).unwrap();
        let p = Project::open(&dir.join("Invoke-AppDeployToolkit.db")).unwrap();
        let issues = validate(&p, &catalog);
        for i in &issues {
            eprintln!("{:?} {} — {}", i.severity, i.location(), i.message);
        }
        assert_eq!(count(&issues, Severity::Error), 0);
    }

    #[test]
    fn validates_project() {
        let Ok(root) = std::env::var("DE_REPO") else { return };
        let paths = crate::paths::AppPaths::with_root(&root);
        let catalog = Catalog::load(&paths.psadt_database()).unwrap();
        let dir = std::env::temp_dir().join("de-validate");
        let _ = std::fs::remove_dir_all(&dir);
        let mut p = Project::create(&paths, &dir).unwrap();
        let dt = DeploymentType::Installation;
        let ok = p.add_action(dt, "Start-ADTProcess", "Run cmd", "", &[("FilePath", "cmd.exe")]).unwrap();
        let missing = p.add_action(dt, "Start-ADTMsiProcess", "MSI", "", &[("FilePath", "nothere.msi")]).unwrap();
        let empty = p.add_action(dt, "New-ADTShortcut", "Shortcut", "", &[("LiteralPath", "x.lnk")]).unwrap();
        let script = p.add_action(dt, crate::catalog::CUSTOM_SCRIPT, "Script", "", &[("Script", "if ($a) {")]).unwrap();
        let num = p.add_action(dt, "Show-ADTInstallationPrompt", "Prompt", "", &[("Message", "Hi"), ("Timeout", "abc")]).unwrap();
        let disabled = p.add_action(dt, "New-ADTFolder", "Disabled", "", &[]).unwrap();
        p.set_disabled(disabled, true).unwrap();
        let mut a = p.action(ok).unwrap().unwrap();
        a.variable_name = "result".into();
        a.conditions = "($x -eq 1".into();
        p.save_action(&a).unwrap();

        let issues = validate(&p, &catalog);
        let for_id = |id: i64| issues.iter().filter(|i| i.action_id == Some(id) && i.severity == Severity::Error).count();
        assert_eq!(for_id(ok), 2, "{issues:#?}"); // variable + condition
        assert_eq!(for_id(missing), 1);
        assert_eq!(for_id(empty), 1); // TargetPath
        assert_eq!(for_id(script), 1);
        assert_eq!(for_id(num), 1);
        assert_eq!(issues.iter().filter(|i| i.action_id == Some(disabled)).count(), 0);
        assert!(issues.iter().any(|i| i.message.contains("uninstall sequence")));
        assert!(issues.iter().any(|i| i.severity == Severity::Hint));
        assert_eq!(issues[0].severity, Severity::Error);

        std::fs::create_dir_all(dir.join("Files")).unwrap();
        std::fs::write(dir.join("Files").join("nothere.msi"), b"x").unwrap();
        let issues = validate(&p, &catalog);
        assert_eq!(issues.iter().filter(|i| i.action_id == Some(missing)).count(), 0);
    }
}
