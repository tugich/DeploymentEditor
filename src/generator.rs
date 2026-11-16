//! Generates the PSADT deployment script from the project database.
//!
//! This is a faithful port of `BuildScript()` / `GenerateDeploymentFile()` from the
//! PureBasic version so that existing projects produce byte-identical output.

use anyhow::{Context, Result};
use std::fs;
use std::path::Path;

use crate::catalog::{Catalog, CUSTOM_SCRIPT};
use crate::project::{keys, Action, DeploymentType, Project};

const CRLF: &str = "\r\n";

/// Build the PowerShell block for one deployment type (the `<InstallationPart>` etc.).
pub fn build_script(actions: &[Action], catalog: &Catalog) -> String {
    let mut out = String::new();
    let mut spacing = "    ";
    let mut block_opened = false;

    for action in actions.iter().filter(|a| !a.disabled) {
        // The original iterated over the LEFT JOIN view: an action without values
        // still yields one row with an empty parameter.
        let empty = [(String::new(), String::new())];
        let rows: &[(String, String)] = if action.values.is_empty() { &empty } else { &action.values };
        let description = action.description.replace(CRLF, "").replace('\n', "");

        for (i, (parameter, value)) in rows.iter().enumerate() {
            let ptype = catalog.parameter_type(&action.command, parameter);
            let mut enclosing = "\"";
            if !ptype.contains("String") {
                enclosing = "";
            }
            if ptype.contains("Guid") {
                enclosing = "\"";
            }

            if action.command == CUSTOM_SCRIPT {
                let normalized = value.replace(CRLF, "\n");
                let body = normalized.replace('\n', &format!("{CRLF}{spacing}"));
                out.push_str(&format!("{CRLF}{spacing}# {description}{CRLF}{spacing}{body}"));
                continue;
            }

            if i == 0 {
                if !description.is_empty() {
                    out.push_str(&format!("{CRLF}{spacing}# {description}"));
                }

                if !action.conditions.is_empty() {
                    if block_opened {
                        out.push_str(&format!("{CRLF}{spacing}}}"));
                    }
                    out.push_str(&format!("{CRLF}{spacing}if ({}) {{ ", action.conditions));
                    spacing = "      ";
                    block_opened = true;
                } else {
                    if block_opened {
                        out.push_str(&format!("{CRLF}{spacing}}}"));
                        block_opened = false;
                    }
                    spacing = "    ";
                }

                let prefix = if action.variable_name.trim().is_empty() {
                    String::new()
                } else {
                    format!("{} = ", action.variable_name)
                };
                out.push_str(&format!("{CRLF}{spacing}{prefix}{}", action.command));
                if action.continue_on_error {
                    out.push_str(" -ErrorAction SilentlyContinue");
                }
            }

            if !value.is_empty() {
                if ptype == "SwitchParameter" {
                    if value == "1" {
                        out.push_str(&format!(" -{parameter}"));
                    }
                } else {
                    out.push_str(&format!(" -{parameter} {enclosing}{value}{enclosing}"));
                }
            }
        }
    }

    if block_opened {
        out.push_str(&format!("{CRLF}    }}"));
    }
    out
}

/// Fill the deployment template with the project settings and generated parts.
pub fn render_template(template: &str, project: &Project, catalog: &Catalog, today: &str) -> Result<String> {
    let install = build_script(&project.actions_for_generation(DeploymentType::Installation)?, catalog);
    let uninstall = build_script(&project.actions_for_generation(DeploymentType::Uninstall)?, catalog);
    let repair = build_script(&project.actions_for_generation(DeploymentType::Repair)?, catalog);

    let mut out = String::with_capacity(template.len() + install.len() + uninstall.len() + repair.len());
    // PureBasic's ReadString() strips CR/LF; WriteString appends CRLF to every line.
    for line in template.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        let line = line
            .replace("<ProductPublisher>", project.setting(keys::APP_VENDOR))
            .replace("<ProductName>", project.setting(keys::APP_NAME))
            .replace("<ProductVersion>", project.setting(keys::APP_VERSION))
            .replace("<ProductLanguage>", project.setting(keys::APP_LANGUAGE))
            .replace("<ProductArchitecture>", project.setting(keys::APP_ARCH))
            .replace("<ScriptDate>", today)
            .replace("<ScriptAuthor>", project.setting(keys::APP_AUTHOR))
            .replace("<InstallationPart>", &install)
            .replace("<UninstallPart>", &uninstall)
            .replace("<RepairPart>", &repair);
        out.push_str(&line);
        out.push_str(CRLF);
    }
    // A template that ends with a newline produced one empty trailing line in PB;
    // `split` yields an empty last element, matching that behaviour except for the
    // extra CRLF which PB did not write (Eof was reached).
    if template.ends_with('\n') {
        out.truncate(out.len() - CRLF.len());
    }
    Ok(out)
}

/// Write `Invoke-AppDeployToolkit.ps1` into the project folder.
/// The per-project `.template` copy is created from the global template if missing.
pub fn generate_deployment_file(project: &Project, catalog: &Catalog, global_template: &Path) -> Result<()> {
    let tpl_path = project.deployment_template_file();
    if !tpl_path.is_file() {
        fs::copy(global_template, &tpl_path)
            .with_context(|| format!("Copying deployment template from {}", global_template.display()))?;
    }
    let bytes = fs::read(&tpl_path).with_context(|| format!("Reading {}", tpl_path.display()))?;
    let template = String::from_utf8_lossy(&bytes);
    let today = chrono::Local::now().format("%Y-%m-%d").to_string();
    let content = render_template(&template, project, catalog, &today)?;
    fs::write(project.deployment_file(), content.as_bytes())
        .with_context(|| format!("Can't write the deployment file {}", project.deployment_file().display()))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::paths::AppPaths;
    use std::path::PathBuf;

    fn repo() -> Option<PathBuf> {
        std::env::var("DE_REPO").ok().map(PathBuf::from).filter(|p| p.join("Databases/PSADT.sqlite").is_file())
    }

    /// The demo project in the original repository ships with a deployment file that
    /// was produced by the PureBasic generator — our output must match it.
    #[test]
    fn matches_original_generator_on_demo_project() {
        let Some(root) = repo() else { eprintln!("DE_REPO not set, skipping"); return };
        let paths = AppPaths::with_root(&root);
        let catalog = Catalog::load(&paths.psadt_database()).unwrap();
        let tmp = std::env::temp_dir().join("de-test-demo");
        let _ = std::fs::remove_dir_all(&tmp);
        crate::project::copy_dir_all(&paths.demo_project(), &tmp).unwrap();
        let project = Project::open(&tmp.join("Invoke-AppDeployToolkit.db")).unwrap();

        let expected = std::fs::read(root.join("Test/Invoke-AppDeployToolkit.ps1")).unwrap();
        // git may have normalised the shipped file to LF; compare with CRLF.
        let expected = crate::plugins::to_crlf(&String::from_utf8_lossy(&expected));
        let date = expected
            .lines()
            .find_map(|l| l.trim().strip_prefix("AppScriptDate = '"))
            .map(|s| s.trim_end_matches('\'').to_string())
            .unwrap();
        let tpl_path = if project.deployment_template_file().is_file() {
            project.deployment_template_file()
        } else {
            paths.template_script()
        };
        let tpl = String::from_utf8_lossy(&std::fs::read(tpl_path).unwrap()).to_string();
        let got = render_template(&tpl, &project, &catalog, &date).unwrap();
        if got != expected {
            std::fs::write(tmp.join("got.ps1"), &got).unwrap();
            let gl: Vec<_> = got.split("\r\n").collect();
            let el: Vec<_> = expected.split("\r\n").collect();
            for (i, (g, e)) in gl.iter().zip(el.iter()).enumerate() {
                if g != e {
                    panic!("line {}:\n got: {:?}\n exp: {:?}", i + 1, g, e);
                }
            }
            panic!("length differs: got {} lines, expected {}", gl.len(), el.len());
        }
    }

    #[test]
    fn conditions_variables_and_switches() {
        let Some(root) = repo() else { return };
        let catalog = Catalog::load(&AppPaths::with_root(&root).psadt_database()).unwrap();
        let a = Action {
            id: 1,
            command: "Start-ADTProcess".into(),
            description: "Run setup".into(),
            variable_name: "$result".into(),
            conditions: "$true".into(),
            continue_on_error: true,
            values: vec![
                ("FilePath".into(), "setup.exe".into()),
                ("NoWait".into(), "1".into()),
                ("UseShellExecute".into(), "0".into()),
            ],
            ..Default::default()
        };
        let b = Action { id: 2, command: CUSTOM_SCRIPT.into(), description: "Hi".into(),
            values: vec![("Script".into(), "Write-Host 1\nWrite-Host 2".into())], ..Default::default() };
        let c = Action { id: 3, command: "Close-ADTInstallationProgress".into(), ..Default::default() };
        let s = build_script(&[a, b, c], &catalog);
        assert_eq!(
            s,
            "\r\n    # Run setup\r\n    if ($true) { \r\n      $result = Start-ADTProcess -ErrorAction SilentlyContinue -FilePath \"setup.exe\" -NoWait\
             \r\n      # Hi\r\n      Write-Host 1\r\n      Write-Host 2\
             \r\n      }\r\n    Close-ADTInstallationProgress"
        );
    }
}
