//! Launching external programs (PSADT, PowerShell, editors, Explorer, ...).
//!
//! Elevated launches use `Start-Process -Verb RunAs` via an *encoded* PowerShell
//! command, which avoids every command-line quoting pitfall and needs no unsafe
//! Win32 FFI.

use anyhow::{anyhow, bail, Result};
use std::path::Path;
use std::process::{Command, Stdio};

#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

/// Single-quote a value for PowerShell (`'` becomes `''`).
pub fn ps_quote(s: &str) -> String {
    format!("'{}'", s.replace('\'', "''"))
}

fn base64(bytes: &[u8]) -> String {
    const T: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(bytes.len().div_ceil(3) * 4);
    for chunk in bytes.chunks(3) {
        let b = [chunk[0], *chunk.get(1).unwrap_or(&0), *chunk.get(2).unwrap_or(&0)];
        let n = ((b[0] as u32) << 16) | ((b[1] as u32) << 8) | b[2] as u32;
        out.push(T[(n >> 18) as usize & 63] as char);
        out.push(T[(n >> 12) as usize & 63] as char);
        out.push(if chunk.len() > 1 { T[(n >> 6) as usize & 63] as char } else { '=' });
        out.push(if chunk.len() > 2 { T[n as usize & 63] as char } else { '=' });
    }
    out
}

/// Encode a script for `powershell.exe -EncodedCommand` (UTF-16LE + Base64).
pub fn encode_ps(script: &str) -> String {
    let bytes: Vec<u8> = script.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
    base64(&bytes)
}

fn powershell_command(script: &str, hidden: bool) -> Command {
    let mut cmd = Command::new("powershell.exe");
    cmd.args(["-NoProfile", "-NoLogo", "-ExecutionPolicy", "Bypass"]);
    if hidden {
        cmd.args(["-WindowStyle", "Hidden"]);
    }
    cmd.args(["-EncodedCommand", &encode_ps(script)]);
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        if hidden {
            cmd.creation_flags(CREATE_NO_WINDOW);
        }
    }
    cmd
}

fn ensure_windows() -> Result<()> {
    if cfg!(windows) {
        Ok(())
    } else {
        bail!("This feature requires Windows.")
    }
}

/// Run a hidden PowerShell script without waiting for it.
pub fn spawn_powershell_hidden(script: &str) -> Result<()> {
    ensure_windows()?;
    powershell_command(script, true).spawn().map(|_| ()).map_err(|e| anyhow!("Can't start PowerShell: {e}"))
}

/// Run a hidden PowerShell script and capture stdout/stderr (blocking).
pub fn run_powershell_capture(script: &str) -> Result<(i32, String)> {
    ensure_windows()?;
    let out = powershell_command(script, true)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|e| anyhow!("Can't start PowerShell: {e}"))?;
    let mut text = String::from_utf8_lossy(&out.stdout).to_string();
    let err = String::from_utf8_lossy(&out.stderr);
    if !err.trim().is_empty() {
        text.push('\n');
        text.push_str(&err);
    }
    Ok((out.status.code().unwrap_or(-1), text))
}

/// Start a program elevated (UAC prompt), like `ShellExecute(..., "RunAs", ...)`.
pub fn run_elevated(file: &Path, args: &str, working_dir: &Path) -> Result<()> {
    let mut script = format!(
        "Start-Process -FilePath {} -WorkingDirectory {} -Verb RunAs",
        ps_quote(&file.to_string_lossy()),
        ps_quote(&working_dir.to_string_lossy())
    );
    if !args.trim().is_empty() {
        script.push_str(&format!(" -ArgumentList {}", ps_quote(args)));
    }
    spawn_powershell_hidden(&script)
}

/// Run a program by name (resolved via PATH, including .cmd shims like `code`).
pub fn run_program(program: &str, file_arg: &Path) -> Result<()> {
    let script = format!(
        "Start-Process -FilePath {} -ArgumentList {}",
        ps_quote(program),
        ps_quote(&format!("\"{}\"", file_arg.to_string_lossy()))
    );
    spawn_powershell_hidden(&script)
}

pub fn open_path(path: &Path) -> Result<()> {
    if !path.exists() {
        bail!("The path does not exist: {}", path.display());
    }
    open::that_detached(path).map_err(|e| anyhow!("Can't open {}: {e}", path.display()))
}

pub fn open_url(url: &str) -> Result<()> {
    open::that_detached(url).map_err(|e| anyhow!("Can't open {url}: {e}"))
}

/// PowerShell console (elevated) in the project folder.
pub fn powershell_here(folder: &Path) -> Result<()> {
    let inner = format!("Set-Location -LiteralPath {}", ps_quote(&folder.to_string_lossy()));
    let args = format!("-NoExit -EncodedCommand {}", encode_ps(&inner));
    run_elevated(Path::new("powershell.exe"), &args, folder)
}

#[derive(Clone, Debug, Default)]
pub struct MsiInfo {
    pub manufacturer: String,
    pub product_name: String,
    pub product_code: String,
    pub product_version: String,
}

/// Read MSI properties with the bundled `Scripts/Read-Msi.ps1` (blocking).
pub fn read_msi(scripts_dir: &Path, msi: &Path) -> Result<MsiInfo> {
    let script = format!(
        "& {} -FilePath {}",
        ps_quote(&scripts_dir.join("Read-Msi.ps1").to_string_lossy()),
        ps_quote(&msi.to_string_lossy())
    );
    let (code, out) = run_powershell_capture(&script)?;
    let mut info = MsiInfo::default();
    for line in out.lines() {
        let line = line.trim();
        if let Some(v) = line.strip_prefix("Productmanufacturer:") {
            info.manufacturer = v.trim().into();
        } else if let Some(v) = line.strip_prefix("Productname:") {
            info.product_name = v.trim().into();
        } else if let Some(v) = line.strip_prefix("Productcode:") {
            info.product_code = v.trim().into();
        } else if let Some(v) = line.strip_prefix("Productversion:") {
            info.product_version = v.trim().into();
        }
    }
    if code != 0 && info.product_code.is_empty() {
        bail!("Reading the MSI failed:\n{out}");
    }
    Ok(info)
}

/// Quote one argument for a Windows command line (CommandLineToArgvW rules):
/// backslashes before a quote — including the closing quote — are doubled.
pub fn win_quote_arg(arg: &str) -> String {
    let mut out = String::from("\"");
    let mut backslashes = 0;
    for c in arg.chars() {
        match c {
            '\\' => backslashes += 1,
            '"' => {
                out.push_str(&"\\".repeat(backslashes * 2 + 1));
                out.push('"');
                backslashes = 0;
            }
            _ => {
                out.push_str(&"\\".repeat(backslashes));
                out.push(c);
                backslashes = 0;
            }
        }
    }
    out.push_str(&"\\".repeat(backslashes * 2));
    out.push('"');
    out
}

/// Project folder exactly as the PureBasic version passed it to plugins: with a
/// trailing backslash (plugins such as Script Validator rely on it).
pub fn plugin_project_path(project_folder: &Path) -> String {
    let mut folder = project_folder.to_string_lossy().trim_end_matches(['\\', '/']).to_string();
    folder.push('\\');
    folder
}

/// Argument list for `powershell.exe` when running a plugin.
pub fn plugin_arguments(script: &Path, project_folder: &Path) -> String {
    format!(
        "-NoLogo -ExecutionPolicy Bypass -File {} -ProjectPath {}",
        win_quote_arg(&script.to_string_lossy()),
        win_quote_arg(&plugin_project_path(project_folder))
    )
}

/// Start `powershell.exe` with its own, visible console window.
///
/// Spawning a console program directly from this GUI program (which has no
/// console) hands it invalid standard handles: the window would flash and close
/// immediately and `Pause` / `Read-Host` would not work. `Start-Process` creates the
/// new console properly, like PureBasic's RunProgram did.
fn start_visible_powershell(arguments: &str, working_dir: &Path) -> Result<()> {
    let script = format!(
        "Start-Process -FilePath 'powershell.exe' -ArgumentList {} -WorkingDirectory {}",
        ps_quote(arguments),
        ps_quote(&working_dir.to_string_lossy())
    );
    spawn_powershell_hidden(&script)
}

/// Run a plugin script in a visible PowerShell window.
pub fn run_plugin(script: &Path, project_folder: &Path) -> Result<()> {
    ensure_windows()?;
    if !script.is_file() {
        bail!("The plugin script was not found: {}", script.display());
    }
    let working_dir = script.parent().unwrap_or(project_folder);
    start_visible_powershell(&plugin_arguments(script, project_folder), working_dir)
        .map_err(|e| anyhow!("Can't start the plugin: {e}"))
}

/// Start the generated "Run on Remote Machine" controller in a visible PowerShell window.
pub fn run_remote_controller(script: &Path) -> Result<()> {
    ensure_windows()?;
    let args = format!("-NoLogo -NoProfile -ExecutionPolicy Bypass -File {}", win_quote_arg(&script.to_string_lossy()));
    start_visible_powershell(&args, script.parent().unwrap_or(script))
        .map_err(|e| anyhow!("Can't start PowerShell: {e}"))
}

/// Interactive PSADT command help console (`Scripts/PSADT4-HelpConsole.ps1`).
pub fn help_console(root: &Path) -> Result<()> {
    ensure_windows()?;
    let script = root.join("Scripts").join("PSADT4-HelpConsole.ps1");
    let args = format!("-NoLogo -ExecutionPolicy Bypass -File {}", win_quote_arg(&script.to_string_lossy()));
    start_visible_powershell(&args, root).map_err(|e| anyhow!("Can't start PowerShell: {e}"))
}

/// Parse PowerShell snippets with the real PowerShell parser (Windows only).
/// Input: (key, code). Output: (key, line, message) for every syntax error.
pub fn powershell_syntax_check(items: &[(i64, String)]) -> Result<Vec<(i64, u32, String)>> {
    ensure_windows()?;
    let base = std::env::temp_dir().join(format!("de-syntax-{}", std::process::id()));
    let input = base.with_extension("in.json");
    let output = base.with_extension("out.json");
    let _ = std::fs::remove_file(&output);
    write_syntax_input(&input, items)?;
    let res = run_powershell_capture(&syntax_check_script(&input, &output));
    let result = std::fs::read(&output);
    let _ = std::fs::remove_file(&input);
    let _ = std::fs::remove_file(&output);
    let (code, console) = res?;
    match result {
        Ok(bytes) => parse_syntax_output(&bytes),
        Err(_) => bail!("PowerShell syntax check failed (exit {code}): {}", console.trim()),
    }
}

fn write_syntax_input(path: &Path, items: &[(i64, String)]) -> Result<()> {
    let input: Vec<serde_json::Value> =
        items.iter().map(|(id, code)| serde_json::json!({ "id": id, "code": code })).collect();
    std::fs::write(path, serde_json::to_vec(&input)?)?;
    Ok(())
}

/// The check writes its result to a file instead of the console: the editor starts
/// PowerShell without a console window, where console encoding calls fail.
fn syntax_check_script(input: &Path, output: &Path) -> String {
    format!(
        "$ErrorActionPreference = 'Stop'\n\
         try {{\n\
           if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {{ throw \"PowerShell runs in $($ExecutionContext.SessionState.LanguageMode) mode (AppLocker / WDAC) - the parser can't be used.\" }}\n\
           $items = Get-Content -Raw -Encoding UTF8 -LiteralPath {input} | ConvertFrom-Json\n\
           $out = @(foreach ($i in @($items)) {{\n\
             $t = $null; $e = $null\n\
             [void][System.Management.Automation.Language.Parser]::ParseInput([string]$i.code, [ref]$t, [ref]$e)\n\
             foreach ($x in @($e)) {{ if ($x) {{ [pscustomobject]@{{ id = [long]$i.id; line = [int]$x.Extent.StartLineNumber; message = [string]$x.Message }} }} }}\n\
           }})\n\
           $json = ConvertTo-Json -InputObject @{{ errors = $out }} -Compress -Depth 4\n\
         }}\n\
         catch {{ $json = ConvertTo-Json -InputObject @{{ failure = [string]$_.Exception.Message }} -Compress }}\n\
         Set-Content -LiteralPath {output} -Value $json -Encoding UTF8\n",
        input = ps_quote(&input.to_string_lossy()),
        output = ps_quote(&output.to_string_lossy()),
    )
}

fn parse_syntax_output(bytes: &[u8]) -> Result<Vec<(i64, u32, String)>> {
    let text = String::from_utf8_lossy(bytes);
    let text = text.trim_start_matches('\u{feff}').trim();
    let v: serde_json::Value =
        serde_json::from_str(text).map_err(|e| anyhow!("Unexpected output of the syntax check: {e}: {text}"))?;
    if let Some(f) = v.get("failure").and_then(|f| f.as_str()) {
        bail!("{f}");
    }
    // Windows PowerShell 5.1 may unwrap a one-element array into an object.
    let list = match v.get("errors") {
        Some(serde_json::Value::Array(a)) => a.clone(),
        Some(serde_json::Value::Null) | None => Vec::new(),
        Some(serde_json::Value::Object(o)) if o.contains_key("value") => {
            o["value"].as_array().cloned().unwrap_or_default()
        }
        Some(other) => vec![other.clone()],
    };
    Ok(list
        .iter()
        .map(|v| {
            (
                v["id"].as_i64().unwrap_or(0),
                v["line"].as_u64().unwrap_or(0) as u32,
                v["message"].as_str().unwrap_or("").to_string(),
            )
        })
        .collect())
}

/// Write `Windows Sandbox.wsb` into the project folder and launch it.
pub fn start_sandbox(template: &Path, project_folder: &Path) -> Result<()> {
    ensure_windows()?;
    let tpl = std::fs::read_to_string(template)
        .map_err(|e| anyhow!("Can't read the Windows Sandbox template {}: {e}", template.display()))?;
    let mut folder = project_folder.to_string_lossy().to_string();
    if !folder.ends_with('\\') {
        folder.push('\\');
    }
    let content: String = tpl
        .lines()
        .map(|l| l.replace("$(ProjectPath)", &folder) + "\r\n")
        .collect();
    let wsb = project_folder.join("Windows Sandbox.wsb");
    std::fs::write(&wsb, content)?;
    open::that_detached(&wsb).map_err(|e| anyhow!("Can't start Windows Sandbox: {e}"))
}

/// Collect `*.exe` names (without extension) below a folder, comma separated.
pub fn executables_list(folder: &Path) -> String {
    fn walk(dir: &Path, out: &mut Vec<String>) {
        let Ok(rd) = std::fs::read_dir(dir) else { return };
        for e in rd.flatten() {
            let p = e.path();
            if p.is_dir() {
                walk(&p, out);
            } else if p.extension().map(|x| x.eq_ignore_ascii_case("exe")).unwrap_or(false) {
                if let Some(stem) = p.file_stem() {
                    out.push(stem.to_string_lossy().to_string());
                }
            }
        }
    }
    let mut v = Vec::new();
    walk(folder, &mut v);
    v.join(",")
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Runs the syntax check script with a real PowerShell when DE_PWSH points to one.
    #[test]
    fn syntax_check_with_real_powershell() {
        let Ok(pwsh) = std::env::var("DE_PWSH") else { return };
        let dir = std::env::temp_dir().join("de-syntax-test");
        std::fs::create_dir_all(&dir).unwrap();
        let cases: Vec<(&str, Vec<(i64, String)>)> = vec![
            ("none", vec![]),
            ("one ok", vec![(1, "Write-Host 'hi'\n".into())]),
            ("one bad", vec![(2, "if ($a {\n  Write-Host 'x'\n".into())]),
            ("mixed", vec![(3, "Get-Process | Out-Null".into()), (4, "Write-Host \"umlaut äöü".into()), (5, "$x = @(1,2".into())]),
        ];
        for (name, items) in cases {
            let input = dir.join(format!("{}.json", name.replace(' ', "_")));
            let output = dir.join(format!("{}.out.json", name.replace(' ', "_")));
            let _ = std::fs::remove_file(&output);
            write_syntax_input(&input, &items).unwrap();
            let out = Command::new(&pwsh)
                .args(["-NoProfile", "-NonInteractive", "-EncodedCommand", &encode_ps(&syntax_check_script(&input, &output))])
                .output()
                .unwrap();
            assert!(out.stdout.is_empty() || String::from_utf8_lossy(&out.stdout).trim().is_empty(), "no console output expected");
            let bytes = std::fs::read(&output).unwrap_or_else(|_| panic!("{name}: no output file; {}", String::from_utf8_lossy(&out.stderr)));
            let res = parse_syntax_output(&bytes);
            eprintln!("{name}: {res:?}\n  raw: {}", String::from_utf8_lossy(&bytes).trim());
            let res = res.unwrap_or_else(|e| panic!("{name}: {e}"));
            let ids: Vec<i64> = res.iter().map(|r| r.0).collect();
            match name {
                "none" | "one ok" => assert!(ids.is_empty()),
                "one bad" => assert!(ids.contains(&2)),
                _ => {
                    assert!(ids.contains(&4) && ids.contains(&5) && !ids.contains(&3));
                }
            }
        }
        // A failing script reports a readable message
        let output = dir.join("broken.out.json");
        let script = syntax_check_script(&dir.join("does-not-exist.json"), &output);
        Command::new(&pwsh).args(["-NoProfile", "-NonInteractive", "-EncodedCommand", &encode_ps(&script)]).output().unwrap();
        let err = parse_syntax_output(&std::fs::read(&output).unwrap()).unwrap_err().to_string();
        eprintln!("broken: {err}");
        assert!(err.contains("does-not-exist"));
    }

    /// Every generated step of the demo project must parse without errors.
    #[test]
    fn demo_steps_parse_with_real_powershell() {
        let (Ok(pwsh), Ok(root)) = (std::env::var("DE_PWSH"), std::env::var("DE_REPO")) else { return };
        let paths = crate::paths::AppPaths::with_root(&root);
        let catalog = crate::catalog::Catalog::load(&paths.psadt_database()).unwrap();
        let dir = std::env::temp_dir().join("de-syntax-demo");
        let _ = std::fs::remove_dir_all(&dir);
        let src = std::env::var("DE_SYNTAX_PROJECT").ok().filter(|s| !s.is_empty()).map(std::path::PathBuf::from).unwrap_or_else(|| paths.demo_project());
        crate::project::copy_dir_all(&src, &dir).unwrap();
        let project = crate::project::Project::open(&dir.join(crate::project::DATABASE_FILE)).unwrap();
        let mut items = Vec::new();
        for dt in crate::project::DeploymentType::ALL {
            for a in project.actions(dt).unwrap() {
                let code = crate::generator::build_script(std::slice::from_ref(&a), &catalog).replace("\r\n", "\n");
                items.push((a.id, code));
            }
        }
        assert!(!items.is_empty());
        eprintln!("checking {} generated steps", items.len());
        let input = dir.join("in.json");
        let output = dir.join("out.json");
        write_syntax_input(&input, &items).unwrap();
        Command::new(&pwsh).args(["-NoProfile", "-NonInteractive", "-EncodedCommand", &encode_ps(&syntax_check_script(&input, &output))]).output().unwrap();
        let res = parse_syntax_output(&std::fs::read(&output).unwrap()).unwrap();
        assert!(res.is_empty(), "unexpected syntax errors: {res:?}");
    }

    #[test]
    fn syntax_output_variants() {
        assert!(parse_syntax_output(b"\xEF\xBB\xBF{\"errors\":[]}").unwrap().is_empty());
        assert!(parse_syntax_output(b"{\"errors\":null}").unwrap().is_empty());
        let one = parse_syntax_output(br#"{"errors":{"id":7,"line":2,"message":"x"}}"#).unwrap();
        assert_eq!(one, vec![(7, 2, "x".to_string())]);
        let wrapped = parse_syntax_output(br#"{"errors":{"value":[{"id":1,"line":1,"message":"a"}],"Count":1}}"#).unwrap();
        assert_eq!(wrapped.len(), 1);
        assert!(parse_syntax_output(br#"{"failure":"Access denied"}"#).unwrap_err().to_string().contains("Access denied"));
    }
    #[test]
    fn base64_works() {
        assert_eq!(base64(b"Man"), "TWFu");
        assert_eq!(base64(b"Ma"), "TWE=");
        assert_eq!(base64(b"M"), "TQ==");
        // PowerShell docs example: 'dir "c:\program files" '
        assert_eq!(encode_ps("dir"), "ZABpAHIA");
    }
    #[test]
    fn plugin_command_line() {
        assert_eq!(win_quote_arg(r"C:\Proj\"), r#""C:\Proj\\""#);
        assert_eq!(win_quote_arg(r"C:\My Apps\x.ps1"), r#""C:\My Apps\x.ps1""#);
        assert_eq!(win_quote_arg(r#"a"b"#), r#""a\"b""#);
        assert_eq!(plugin_project_path(Path::new(r"C:\Pkg\7zip")), r"C:\Pkg\7zip\");
        assert_eq!(plugin_project_path(Path::new(r"C:\Pkg\7zip\")), r"C:\Pkg\7zip\");
        assert_eq!(
            plugin_arguments(Path::new(r"C:\DE\Plugins\Enabled-HelloWorld\HelloWorld.ps1"), Path::new(r"C:\Pkg\7 Zip")),
            r#"-NoLogo -ExecutionPolicy Bypass -File "C:\DE\Plugins\Enabled-HelloWorld\HelloWorld.ps1" -ProjectPath "C:\Pkg\7 Zip\\""#
        );
    }

    #[test]
    fn quoting() {
        assert_eq!(ps_quote("O'Neil"), "'O''Neil'");
    }
}
