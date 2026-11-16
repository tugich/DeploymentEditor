//! PDF documentation of a package: what it installs, how Intune detects it, and
//! every step of the installation / uninstall / repair procedure.

use anyhow::Result;

use crate::catalog::Catalog;
use crate::detection::DetectionRule;
use crate::generator;
use crate::pdf::{self, Font, Pdf, Rgb, CONTENT_W, MARGIN_L, PAGE_W};
use crate::project::{keys, Action, DeploymentType, Project};
use crate::ui::intune_upload::IntuneAppSettings;
use crate::validate::{Issue, Severity};

const BRAND: Rgb = Rgb::hex(0x157EFB);
const BRAND_DARK: Rgb = Rgb::hex(0x0B5FC6);
const TEXT: Rgb = Rgb::hex(0x1F2328);
const MUTED: Rgb = Rgb::hex(0x5B636D);
const FAINT: Rgb = Rgb::hex(0x9AA1AA);
const CODE_BG: Rgb = Rgb::hex(0xF3F5F8);
const RED: Rgb = Rgb::hex(0xC42B1C);
const AMBER: Rgb = Rgb::hex(0xB46A00);
const WHITE: Rgb = Rgb(1.0, 1.0, 1.0);

#[derive(Clone, Copy, Debug)]
pub struct ReportOptions {
    pub include_disabled: bool,
    pub include_code: bool,
    pub include_checks: bool,
}

impl Default for ReportOptions {
    fn default() -> Self {
        Self { include_disabled: true, include_code: true, include_checks: true }
    }
}

fn section_title(pdf: &mut Pdf, title: &str, subtitle: Option<&str>) {
    pdf.ensure(90.0);
    pdf.space(6.0);
    pdf.bookmark(title, 0);
    pdf.rect(MARGIN_L, pdf.y + 2.0, 3.0, 15.0, BRAND);
    pdf.text_at(MARGIN_L + 10.0, pdf.y, Font::Bold, 14.0, TEXT, title);
    pdf.y += 21.0;
    if let Some(s) = subtitle {
        pdf.paragraph(s, Font::Regular, 9.0, MUTED, MARGIN_L, CONTENT_W);
    }
    pdf.space(4.0);
}

fn param_rows(a: &Action, catalog: &Catalog) -> Vec<(String, String)> {
    a.values
        .iter()
        .filter(|(_, v)| !v.trim().is_empty())
        .filter_map(|(k, v)| {
            let info = catalog.parameter(&a.command, k);
            let switch = info.map(|i| i.ptype == "SwitchParameter" || i.is_checkbox()).unwrap_or(false);
            if switch {
                (v == "1").then(|| (format!("-{k}"), "Yes".to_string()))
            } else {
                Some((format!("-{k}"), v.trim().to_string()))
            }
        })
        .collect()
}

fn step(pdf: &mut Pdf, n: usize, a: &Action, catalog: &Catalog, issues: &[Issue]) {
    let indent = MARGIN_L + 28.0;
    let width = CONTENT_W - 28.0;
    pdf.ensure(64.0);
    let top = pdf.y;
    let badge = if a.disabled { FAINT } else { BRAND };
    pdf.circle(MARGIN_L + 10.0, top + 8.0, 10.0, badge);
    let num = n.to_string();
    let w = pdf::text_width(&num, Font::Bold, 9.0);
    pdf.text_at(MARGIN_L + 10.0 - w / 2.0, top + 4.2, Font::Bold, 9.0, WHITE, &num);

    let name = if a.name.trim().is_empty() { "(unnamed step)" } else { a.name.trim() };
    let name_color = if a.disabled { FAINT } else { TEXT };
    pdf.paragraph(name, Font::Bold, 11.5, name_color, indent, width);
    pdf.y -= 2.0;
    let cmd = if a.is_custom_script() { "Custom PowerShell script".to_string() } else { a.command.clone() };
    pdf.paragraph(&cmd, Font::Mono, 8.5, MUTED, indent, width);
    pdf.space(2.0);

    if a.disabled {
        pdf.paragraph("Disabled - this step is not executed.", Font::Italic, 9.0, AMBER, indent, width);
    }
    if !a.description.trim().is_empty() {
        pdf.paragraph(a.description.trim(), Font::Regular, 9.5, TEXT, indent, width);
    }
    let mut notes: Vec<String> = Vec::new();
    if !a.conditions.trim().is_empty() {
        pdf.paragraph(
            &format!("Runs only if: {}", crate::condition::describe(&a.conditions)),
            Font::Bold,
            9.0,
            Rgb::hex(0x6B4FD8),
            indent,
            width,
        );
        pdf.paragraph(&format!("if ({})", a.conditions.trim()), Font::Mono, 8.0, MUTED, indent + 8.0, width - 8.0);
    }
    if !a.variable_name.trim().is_empty() {
        notes.push(format!("Result is stored in {}", a.variable_name.trim()));
    }
    if a.continue_on_error {
        notes.push("Continues on error (-ErrorAction SilentlyContinue)".into());
    }
    for note in notes {
        pdf.paragraph(&format!("- {note}"), Font::Regular, 9.0, MUTED, indent, width);
    }

    if a.is_custom_script() {
        pdf.space(3.0);
        pdf.code_block(a.value("Script"), indent, width, 8.0, CODE_BG, TEXT);
    } else {
        let rows = param_rows(a, catalog);
        if !rows.is_empty() {
            pdf.space(3.0);
            let key_w = rows.iter().map(|(k, _)| pdf::text_width(k, Font::Mono, 8.5)).fold(0.0, f32::max).clamp(60.0, 150.0) + 20.0;
            for (k, v) in rows {
                let lines = pdf::wrap(&v, Font::Mono, 8.5, width - key_w - 8.0);
                pdf.ensure(13.0);
                let row_top = pdf.y;
                pdf.text_at(indent + 6.0, row_top + 2.0, Font::Mono, 8.5, MUTED, &k);
                for l in lines {
                    pdf.ensure(12.0);
                    pdf.text_at(indent + key_w, pdf.y + 2.0, Font::Mono, 8.5, TEXT, &l);
                    pdf.y += 11.5;
                }
                pdf.y += 2.0;
            }
        }
    }
    for i in issues.iter().filter(|i| i.action_id == Some(a.id) && i.severity != Severity::Hint) {
        let (label, color) = if i.severity == Severity::Error { ("Error", RED) } else { ("Warning", AMBER) };
        pdf.paragraph(&format!("{label}: {}", i.message), Font::Italic, 8.5, color, indent, width);
    }
    pdf.space(8.0);
    pdf.hline(indent, MARGIN_L + CONTENT_W, pdf.y - 3.0, 0.5, Rgb::hex(0xE3E5E8));
    pdf.space(6.0);
}

pub fn build(project: &Project, catalog: &Catalog, issues: &[Issue], opts: ReportOptions) -> Result<Vec<u8>> {
    let mut pdf = Pdf::new();
    pdf.muted = MUTED;
    let app = project.setting(keys::APP_NAME).trim().to_string();
    let version = project.setting(keys::APP_VERSION).trim().to_string();
    let vendor = project.setting(keys::APP_VENDOR).trim().to_string();
    let title_name = if app.is_empty() { project.display_name() } else { app.clone() };
    let full_title = format!("{title_name} {version}").trim().to_string();
    pdf.title = format!("{full_title} - Installation procedure");
    pdf.author = project.setting(keys::APP_AUTHOR).to_string();
    pdf.footer_left = format!("{full_title}  |  Installation procedure  |  Deployment Editor");
    let now = chrono::Local::now().format("%Y-%m-%d %H:%M").to_string();

    // ---------------------------------------------------------------- cover band
    let band_h = 128.0;
    pdf.rect(0.0, 0.0, PAGE_W, band_h, BRAND);
    pdf.rect(0.0, band_h, PAGE_W, 4.0, BRAND_DARK);
    pdf.text_at(MARGIN_L, 34.0, Font::Bold, 9.0, Rgb::hex(0xD6E6FF), "INSTALLATION PROCEDURE");
    pdf.text_at(MARGIN_L, 50.0, Font::Bold, 24.0, WHITE, &full_title);
    let sub = [vendor.as_str(), project.setting(keys::APP_ARCH).trim(), project.setting(keys::APP_LANGUAGE).trim()]
        .iter()
        .filter(|s| !s.is_empty())
        .cloned()
        .collect::<Vec<_>>()
        .join("  ·  ");
    pdf.text_at(MARGIN_L, 82.0, Font::Regular, 11.0, WHITE, &sub);
    pdf.text_at(MARGIN_L, 102.0, Font::Regular, 8.5, Rgb::hex(0xD6E6FF), &format!("PSAppDeployToolkit package  ·  documented {now}"));
    let icon = project.folder.join("Assets").join("AppIcon.png");
    if let Some(img) = std::fs::read(&icon).ok().and_then(|b| pdf::Image::from_png(&b)) {
        let idx = pdf.add_image(img);
        pdf.rounded_rect(PAGE_W - MARGIN_L - 76.0, 26.0, 76.0, 76.0, 14.0, WHITE);
        pdf.image(idx, PAGE_W - MARGIN_L - 66.0, 36.0, 56.0, 56.0);
    }
    pdf.y = band_h + 30.0;

    // ---------------------------------------------------------------- overview
    section_title(&mut pdf, "Package overview", None);
    let counts: Vec<String> = DeploymentType::ALL
        .iter()
        .map(|dt| {
            let all = project.actions(*dt).unwrap_or_default();
            let active = all.iter().filter(|a| !a.disabled).count();
            format!("{} {}", dt.as_db(), active)
        })
        .collect();
    let rows: Vec<(String, String, Font)> = vec![
        ("Application".into(), app.clone(), Font::Bold),
        ("Version".into(), version.clone(), Font::Regular),
        ("Vendor".into(), vendor.clone(), Font::Regular),
        ("Architecture".into(), project.setting(keys::APP_ARCH).to_string(), Font::Regular),
        ("Language".into(), project.setting(keys::APP_LANGUAGE).to_string(), Font::Regular),
        ("Script author".into(), project.setting(keys::APP_AUTHOR).to_string(), Font::Regular),
        ("Active steps".into(), counts.join("   ·   "), Font::Regular),
        ("Project folder".into(), project.folder.display().to_string(), Font::Mono),
        ("Deployment file".into(), crate::project::DEPLOYMENT_FILE.into(), Font::Mono),
    ];
    pdf.key_values(&rows, 120.0, 9.5, TEXT);
    pdf.space(10.0);

    // ---------------------------------------------------------------- Intune
    let intune = IntuneAppSettings::for_project(project);
    section_title(&mut pdf, "Deployment with Microsoft Intune", Some("Win32 app settings used when the package is uploaded to Intune."));
    let restart = match intune.restart_behavior.as_str() {
        "basedOnReturnCode" => "Determine behavior based on return codes",
        "allow" => "App install may force a device restart",
        "force" => "Intune will force a mandatory device restart",
        _ => "No specific action",
    };
    let min_os = crate::intune::MinOs::from_key(&intune.min_os).map(|m| m.label()).unwrap_or("-");
    let rows: Vec<(String, String, Font)> = vec![
        ("Install command".into(), intune.install_cmd.clone(), Font::Mono),
        ("Uninstall command".into(), intune.uninstall_cmd.clone(), Font::Mono),
        ("Install behavior".into(), if intune.run_as_system { "System".into() } else { "User".into() }, Font::Regular),
        ("Device restart".into(), restart.into(), Font::Regular),
        ("Requirements".into(), format!("{}  ·  minimum {min_os}", intune.architectures.replace(',', " + ")), Font::Regular),
        ("Return codes".into(), "0, 1707 success  ·  3010 soft reboot  ·  1641 hard reboot  ·  1618 retry".into(), Font::Regular),
        ("Intune app ID".into(), project.setting("Intune_AppId").to_string(), Font::Mono),
    ];
    pdf.key_values(&rows, 120.0, 9.5, TEXT);
    pdf.space(8.0);
    pdf.ensure(40.0);
    pdf.text_at(MARGIN_L, pdf.y, Font::Bold, 10.5, TEXT, "Detection rule");
    pdf.y += 16.0;
    match DetectionRule::load(project) {
        Some(rule) => {
            let kind = match rule.kind() {
                crate::detection::RuleKind::Msi => "MSI product code",
                crate::detection::RuleKind::File => "File system",
                crate::detection::RuleKind::Registry => "Registry",
                crate::detection::RuleKind::Script => "Custom script",
            };
            pdf.paragraph(&format!("{kind}: {}", rule.summary()), Font::Regular, 9.5, TEXT, MARGIN_L, CONTENT_W);
            if let DetectionRule::Script { script, .. } = &rule {
                pdf.space(3.0);
                pdf.code_block(script, MARGIN_L, CONTENT_W, 7.8, CODE_BG, TEXT);
            }
            for p in rule.problems() {
                pdf.paragraph(&format!("Warning: {p}"), Font::Italic, 8.5, AMBER, MARGIN_L, CONTENT_W);
            }
        }
        None => pdf.paragraph("No detection rule has been configured yet.", Font::Italic, 9.5, AMBER, MARGIN_L, CONTENT_W),
    }
    pdf.space(10.0);

    // ---------------------------------------------------------------- checks
    if opts.include_checks {
        let errors = issues.iter().filter(|i| i.severity == Severity::Error).count();
        let warnings = issues.iter().filter(|i| i.severity == Severity::Warning).count();
        section_title(&mut pdf, "Quality check", None);
        pdf.ensure(30.0);
        let (bg, fg, text) = if errors > 0 {
            (Rgb::hex(0xFBE9E7), RED, format!("{errors} error(s) and {warnings} warning(s) found when this document was created."))
        } else if warnings > 0 {
            (Rgb::hex(0xFFF4E0), AMBER, format!("No errors, {warnings} warning(s) found when this document was created."))
        } else {
            (Rgb::hex(0xE7F6EC), Rgb::hex(0x107C41), "No problems found when this document was created.".to_string())
        };
        pdf.rounded_rect(MARGIN_L, pdf.y, CONTENT_W, 24.0, 6.0, bg);
        pdf.text_at(MARGIN_L + 10.0, pdf.y + 7.5, Font::Bold, 9.5, fg, &text);
        pdf.y += 32.0;
        for i in issues.iter().filter(|i| i.severity != Severity::Hint) {
            let color = if i.severity == Severity::Error { RED } else { AMBER };
            pdf.paragraph(&format!("{}: {}", i.location(), i.message), Font::Regular, 8.8, color, MARGIN_L, CONTENT_W);
        }
        pdf.space(8.0);
    }

    // ---------------------------------------------------------------- procedures
    for dt in DeploymentType::ALL {
        let all = project.actions(dt)?;
        let list: Vec<&Action> = all.iter().filter(|a| opts.include_disabled || !a.disabled).collect();
        let (title, intro) = match dt {
            DeploymentType::Installation => ("Installation procedure", "Steps executed by: Invoke-AppDeployToolkit.exe -DeploymentType Install"),
            DeploymentType::Uninstall => ("Uninstall procedure", "Steps executed by: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall"),
            DeploymentType::Repair => ("Repair procedure", "Steps executed by: Invoke-AppDeployToolkit.exe -DeploymentType Repair"),
        };
        if list.is_empty() && dt == DeploymentType::Repair {
            continue;
        }
        if dt == DeploymentType::Installation || pdf.y > 520.0 {
            pdf.new_page();
        }
        section_title(&mut pdf, title, Some(intro));
        if list.is_empty() {
            pdf.paragraph("No steps defined.", Font::Italic, 9.5, MUTED, MARGIN_L, CONTENT_W);
            continue;
        }
        let mut n = 0;
        for a in list {
            n += 1;
            step(&mut pdf, n, a, catalog, issues);
        }
    }

    // ---------------------------------------------------------------- appendix
    if opts.include_code {
        pdf.new_page();
        section_title(
            &mut pdf,
            "Appendix: generated PowerShell",
            Some("The code Deployment Editor writes into Invoke-AppDeployToolkit.ps1 for each deployment type."),
        );
        for dt in DeploymentType::ALL {
            let actions = project.actions_for_generation(dt)?;
            if actions.is_empty() {
                continue;
            }
            let code = generator::build_script(&actions, catalog);
            let code: String = code
                .trim_start_matches("\r\n")
                .split("\r\n")
                .map(|l| l.strip_prefix("    ").unwrap_or(l))
                .collect::<Vec<_>>()
                .join("\n");
            pdf.ensure(60.0);
            pdf.bookmark(dt.as_db(), 1);
            pdf.text_at(MARGIN_L, pdf.y, Font::Bold, 10.5, TEXT, dt.as_db());
            pdf.y += 16.0;
            pdf.code_block(&code, MARGIN_L, CONTENT_W, 7.6, CODE_BG, TEXT);
            pdf.space(12.0);
        }
    }

    Ok(pdf.finish())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::paths::AppPaths;

    #[test]
    fn demo_project_report() {
        let Ok(root) = std::env::var("DE_REPO") else { return };
        let paths = AppPaths::with_root(&root);
        let catalog = Catalog::load(&paths.psadt_database()).unwrap();
        let dir = match std::env::var("DE_REPORT_PROJECT") {
            Ok(d) => std::path::PathBuf::from(d),
            Err(_) => {
                let dir = std::env::temp_dir().join("de-report-demo");
                let _ = std::fs::remove_dir_all(&dir);
                crate::project::copy_dir_all(&paths.demo_project(), &dir).unwrap();
                dir
            }
        };
        let project = Project::open(&dir.join(crate::project::DATABASE_FILE)).unwrap();
        let issues = crate::validate::validate(&project, &catalog);
        let bytes = build(&project, &catalog, &issues, ReportOptions::default()).unwrap();
        assert!(bytes.starts_with(b"%PDF-1."));
        if let Ok(out) = std::env::var("DE_REPORT_OUT") {
            std::fs::write(out, &bytes).unwrap();
        }
    }
}
