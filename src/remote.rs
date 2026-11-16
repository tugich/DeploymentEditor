//! "Run on Remote Machine": copy the package to one or more Windows computers and
//! start the deployment there.
//!
//! The editor zips the package and writes two scripts:
//! * `DE-Run.ps1` runs on the target: unpacks the package, starts
//!   `Invoke-AppDeployToolkit.exe` and writes the exit code to `result.txt`.
//! * the controller runs locally in a visible PowerShell window: connects,
//!   copies, runs, waits, copies the PSADT logs back, cleans up and prints a summary.
//!
//! Two ways to reach the computer:
//! * PowerShell Remoting (WinRM) — runs as the connecting admin account.
//! * Scheduled task as SYSTEM — copies through the admin share (`\\pc\C$`) and
//!   registers a one-off task over CIM/DCOM (no WinRM needed). Same context as
//!   Intune; PSADT 4.1 shows its dialogs to the logged-on user in this mode.

use std::path::{Path, PathBuf};

use anyhow::{bail, Context, Result};
use serde::{Deserialize, Serialize};

use crate::project::DeploymentType;
use crate::system::ps_quote;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum Method {
    Auto,
    WinRm,
    Task,
}

impl Method {
    pub fn ps(&self) -> &'static str {
        match self {
            Method::Auto => "Auto",
            Method::WinRm => "WinRM",
            Method::Task => "Task",
        }
    }
    pub fn label(&self) -> &'static str {
        match self {
            Method::Auto => "Automatic",
            Method::WinRm => "PowerShell Remoting",
            Method::Task => "Scheduled task (SYSTEM)",
        }
    }
    pub fn help(&self) -> &'static str {
        match self {
            Method::Auto => "Uses PowerShell Remoting when WinRM answers, otherwise a scheduled task. Interactive mode always uses the scheduled task.",
            Method::WinRm => "Copies and runs through a PowerShell Remoting session as your admin account. Needs WinRM on the target (Enable-PSRemoting). No dialogs can be shown.",
            Method::Task => "Copies through the admin share \\\\computer\\C$ and runs a one-off scheduled task as SYSTEM — the same context as Intune. Needs file sharing and remote task management (RPC) in the firewall.",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct RemoteOptions {
    /// Computer names, separated by new lines, commas or spaces.
    pub computers: String,
    pub method: Method,
    #[serde(skip)]
    pub deployment: Option<DeploymentType>,
    pub deploy_mode: String,
    pub ask_credentials: bool,
    pub remote_dir: String,
    pub cleanup: bool,
    pub copy_logs: bool,
    pub timeout_minutes: u32,
}

impl Default for RemoteOptions {
    fn default() -> Self {
        Self {
            computers: String::new(),
            method: Method::Auto,
            deployment: Some(DeploymentType::Installation),
            deploy_mode: "Silent".into(),
            ask_credentials: false,
            remote_dir: "C:\\Windows\\Temp\\DeploymentEditor".into(),
            cleanup: true,
            copy_logs: true,
            timeout_minutes: 60,
        }
    }
}

/// Split and check the computer list.
pub fn computer_list(text: &str) -> Result<Vec<String>> {
    let mut out: Vec<String> = Vec::new();
    for name in text.split(|c: char| c == ',' || c == ';' || c.is_whitespace()).map(str::trim).filter(|s| !s.is_empty()) {
        let ok = name.len() <= 253
            && name.chars().all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '.' || c == '_')
            && !name.starts_with('-')
            && !name.starts_with('.');
        if !ok {
            bail!("\"{name}\" is not a valid computer name.");
        }
        if !out.iter().any(|n| n.eq_ignore_ascii_case(name)) {
            out.push(name.to_string());
        }
    }
    if out.is_empty() {
        bail!("Enter at least one computer name.");
    }
    Ok(out)
}

/// Check the remote folder: absolute local path on the target, e.g. C:\Windows\Temp\X.
pub fn check_remote_dir(dir: &str) -> Result<String> {
    let d = dir.trim().trim_end_matches('\\');
    let b = d.as_bytes();
    let ok = b.len() >= 4
        && b[0].is_ascii_alphabetic()
        && b[1] == b':'
        && b[2] == b'\\'
        && !d.contains("..")
        && !d.chars().any(|c| matches!(c, '"' | '*' | '?' | '<' | '>' | '|' | '$' | '`') || c.is_control());
    if !ok {
        bail!("The remote folder must be a local folder on the target, e.g. C:\\Windows\\Temp\\DeploymentEditor.");
    }
    if d.matches('\\').count() < 2 {
        bail!("Use a sub folder (e.g. C:\\Windows\\Temp\\DeploymentEditor) — the folder is deleted after the deployment.");
    }
    Ok(d.to_string())
}

fn safe_name(s: &str) -> String {
    let n: String = s
        .chars()
        .map(|c| if c.is_ascii_alphanumeric() || c == '-' || c == '_' || c == '.' { c } else { '_' })
        .collect();
    let n = n.trim_matches('_').to_string();
    if n.is_empty() {
        "Package".into()
    } else {
        n.chars().take(60).collect()
    }
}

/// Script executed on the target computer.
pub fn runner_script(dt: DeploymentType, deploy_mode: &str) -> String {
    let args = format!("-DeploymentType {} -DeployMode {}", dt.as_psadt(), deploy_mode);
    format!(
        r#"# Generated by Deployment Editor - runs on the target computer
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pkg = Join-Path $dir 'Package'
$zip = Join-Path $dir 'package.zip'
$code = 60001
try {{
    if (Test-Path -LiteralPath $pkg) {{ Remove-Item -LiteralPath $pkg -Recurse -Force }}
    Expand-Archive -LiteralPath $zip -DestinationPath $pkg -Force
    Remove-Item -LiteralPath $zip -Force
    Get-ChildItem -LiteralPath $pkg -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
    $p = Start-Process -FilePath (Join-Path $pkg 'Invoke-AppDeployToolkit.exe') -ArgumentList {args} -WorkingDirectory $pkg -Wait -PassThru
    $code = $p.ExitCode
}}
catch {{
    ($_ | Out-String) | Set-Content -LiteralPath (Join-Path $dir 'error.txt') -Encoding UTF8
}}
Set-Content -LiteralPath (Join-Path $dir 'result.tmp') -Value $code -Encoding ASCII
Move-Item -LiteralPath (Join-Path $dir 'result.tmp') -Destination (Join-Path $dir 'result.txt') -Force
exit $code
"#,
        args = ps_quote(&args)
    )
}

pub struct ControllerInput<'a> {
    pub title: &'a str,
    pub computers: &'a [String],
    pub opts: &'a RemoteOptions,
    pub package_zip: &'a Path,
    pub runner: &'a Path,
    pub local_logs: &'a Path,
    pub project_key: &'a str,
}

/// Script executed locally in a visible PowerShell window.
pub fn controller_script(i: &ControllerInput) -> String {
    let dt = i.opts.deployment.unwrap_or(DeploymentType::Installation);
    let list = i.computers.iter().map(|c| ps_quote(c)).collect::<Vec<_>>().join(", ");
    let remote_dir = format!("{}\\{}", i.opts.remote_dir.trim().trim_end_matches('\\'), i.project_key);
    let label = format!("{} ({})", dt.as_psadt(), i.opts.deploy_mode);
    let header = format!(
        r#"# Generated by Deployment Editor - Run on Remote Machine
# {title}
$ErrorActionPreference = 'Stop'
try {{ $Host.UI.RawUI.WindowTitle = {window_title} }} catch {{ }}

$Computers  = @({list})
$Method     = {method}
$DeployMode = {mode}
$Label      = {label}
$Package    = {package}
$Runner     = {runner}
$RemoteDir  = {remote_dir}
$LocalLogs  = {local_logs}
$CopyLogs   = ${copy_logs}
$Cleanup    = ${cleanup}
$TimeoutMin = {timeout}
$TaskName   = {task_name}
$TaskPath   = '\DeploymentEditor\'
$AskCredential = ${ask}
"#,
        title = i.title.replace(['\r', '\n'], " "),
        window_title = ps_quote(&format!("Deployment Editor - Remote {label}: {}", i.title)),
        method = ps_quote(i.opts.method.ps()),
        mode = ps_quote(&i.opts.deploy_mode),
        label = ps_quote(&label),
        package = ps_quote(&i.package_zip.to_string_lossy()),
        runner = ps_quote(&i.runner.to_string_lossy()),
        remote_dir = ps_quote(&remote_dir),
        local_logs = ps_quote(&i.local_logs.to_string_lossy()),
        copy_logs = i.opts.copy_logs,
        cleanup = i.opts.cleanup,
        timeout = i.opts.timeout_minutes.clamp(1, 1440),
        task_name = ps_quote(&format!("Deploy {}", i.project_key)),
        ask = i.opts.ask_credentials,
    );
    header + CONTROLLER_BODY
}

const CONTROLLER_BODY: &str = r#"
function Write-Step([string]$Text) { Write-Host "   $Text" -ForegroundColor Gray }
function Write-Good([string]$Text) { Write-Host "   $Text" -ForegroundColor Green }
function Write-Bad([string]$Text) { Write-Host "   $Text" -ForegroundColor Red }

function Get-ExitText([int]$Code) {
    switch ($Code) {
        0 { return 'Success' }
        1707 { return 'Success' }
        3010 { return 'Success - restart required' }
        1641 { return 'Success - restart initiated' }
        1618 { return 'Another installation is in progress - try again later' }
        1602 { return 'Cancelled by the user' }
    }
    if ($Code -ge 60000 -and $Code -lt 69000) { return 'PSAppDeployToolkit error - see the log' }
    if ($Code -ge 69000 -and $Code -lt 70000) { return 'Error in the deployment script - see the log' }
    return 'Failed - see the log'
}

function Test-Success([int]$Code) { return $Code -in 0, 1707, 3010, 1641 }

function Get-Hint([string]$Message) {
    if ($Message -match 'Access is denied|Zugriff verweigert|0x80070005') { return 'Use an account with administrator rights on the target ("Ask for credentials").' }
    if ($Message -match 'WinRM|WS-Management') { return 'Enable PowerShell Remoting on the target (Enable-PSRemoting) or use the scheduled task method.' }
    if ($Message -match 'RPC server|RPC-Server|0x800706BA') { return 'The target is offline or the firewall blocks remote management (RPC). Allow "Remote Scheduled Tasks Management" and "Windows Management Instrumentation", or use PowerShell Remoting.' }
    if ($Message -match 'network path|Netzwerkpfad|BadNetPath') { return 'The admin share is not reachable. Allow "File and Printer Sharing" in the firewall on the target.' }
    return $null
}

$Drives = @{}
function Get-RemoteRoot([string]$Computer, [string]$Share) {
    $unc = "\\$Computer\$Share"
    if (-not $script:Credential) { return $unc }
    if (-not $Drives.ContainsKey($unc)) {
        $name = 'DE' + [guid]::NewGuid().ToString('N').Substring(0, 8)
        New-PSDrive -Name $name -PSProvider FileSystem -Root $unc -Credential $script:Credential -Scope Script | Out-Null
        $Drives[$unc] = "${name}:"
    }
    return $Drives[$unc]
}

function Get-RemoteLogs([string]$Computer, [datetime]$Since, $Session) {
    if (-not $CopyLogs) { return }
    $target = Join-Path $LocalLogs $Computer
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    $count = 0
    if ($Session) {
        $files = Invoke-Command -Session $Session -ScriptBlock {
            param($s)
            foreach ($d in @("$env:WinDir\Logs\Software", "$env:ProgramData\Logs\Software")) {
                Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $s } | ForEach-Object { $_.FullName }
            }
        } -ArgumentList $Since
        foreach ($f in @($files)) { Copy-Item -FromSession $Session -LiteralPath $f -Destination $target -Force; $count++ }
    }
    else {
        $root = Get-RemoteRoot $Computer 'ADMIN$'
        Get-ChildItem -LiteralPath (Join-Path $root 'Logs\Software') -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $Since } |
            ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $target -Force; $count++ }
    }
    if ($count -gt 0) { Write-Step "$count log file(s) copied to $target" } else { Write-Step 'No new log files found on the target.' }
}

function Invoke-ViaWinRM([string]$Computer) {
    $credArgs = @{}
    if ($script:Credential) { $credArgs.Credential = $script:Credential }
    Write-Step 'Connecting with PowerShell Remoting ...'
    $s = New-PSSession -ComputerName $Computer @credArgs
    try {
        $since = Invoke-Command -Session $s -ScriptBlock { Get-Date }
        Invoke-Command -Session $s -ScriptBlock {
            param($d)
            if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force }
            New-Item -ItemType Directory -Path $d -Force | Out-Null
        } -ArgumentList $RemoteDir
        Write-Step ("Copying package ({0:N1} MB) ..." -f ((Get-Item -LiteralPath $Package).Length / 1MB))
        Copy-Item -LiteralPath $Package -Destination "$RemoteDir\package.zip" -ToSession $s
        Copy-Item -LiteralPath $Runner -Destination "$RemoteDir\DE-Run.ps1" -ToSession $s
        Write-Step "Running $Label as $((Invoke-Command -Session $s -ScriptBlock { "$env:USERDOMAIN\$env:USERNAME" })) ..."
        $out = Invoke-Command -Session $s -ScriptBlock {
            param($d)
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $d 'DE-Run.ps1') | Out-Null
            $LASTEXITCODE
        } -ArgumentList $RemoteDir
        $code = [int]($out | Select-Object -Last 1)
        $err = Invoke-Command -Session $s -ScriptBlock { param($d) $f = Join-Path $d 'error.txt'; if (Test-Path -LiteralPath $f) { Get-Content -LiteralPath $f -Raw } } -ArgumentList $RemoteDir
        if ($err) { Write-Bad $err.Trim() }
        Get-RemoteLogs $Computer $since $s
        if ($Cleanup) {
            Invoke-Command -Session $s -ScriptBlock { param($d) Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue } -ArgumentList $RemoteDir
        }
        return $code
    }
    finally { Remove-PSSession $s -ErrorAction SilentlyContinue }
}

function Invoke-ViaTask([string]$Computer) {
    $credArgs = @{}
    if ($script:Credential) { $credArgs.Credential = $script:Credential }
    $root = Get-RemoteRoot $Computer ($RemoteDir.Substring(0, 1) + '$')
    $base = $root + $RemoteDir.Substring(2)
    Write-Step "Copying package to $base ..."
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
    New-Item -ItemType Directory -Path $base -Force | Out-Null
    Copy-Item -LiteralPath $Package -Destination (Join-Path $base 'package.zip')
    Copy-Item -LiteralPath $Runner -Destination (Join-Path $base 'DE-Run.ps1')

    Write-Step 'Registering scheduled task (SYSTEM) ...'
    $cim = New-CimSession -ComputerName $Computer -SessionOption (New-CimSessionOption -Protocol Dcom) @credArgs
    try {
        $since = (Get-CimInstance -CimSession $cim -ClassName Win32_OperatingSystem).LocalDateTime
        $runner = "$RemoteDir\DE-Run.ps1"
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runner`"" -WorkingDirectory $RemoteDir
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes ($TimeoutMin + 5))
        Register-ScheduledTask -CimSession $cim -TaskName $TaskName -TaskPath $TaskPath -Action $action -Principal $principal -Settings $settings -Force | Out-Null
        Start-ScheduledTask -CimSession $cim -TaskName $TaskName -TaskPath $TaskPath
        $result = Join-Path $base 'result.txt'
        $started = Get-Date
        $deadline = $started.AddMinutes($TimeoutMin)
        while (-not (Test-Path -LiteralPath $result)) {
            if ((Get-Date) -gt $deadline) { throw "No result after $TimeoutMin minutes - the deployment may still be running on $Computer." }
            $elapsed = (Get-Date) - $started
            Write-Host ("`r   Running $Label as SYSTEM ... {0:mm\:ss}" -f $elapsed) -NoNewline -ForegroundColor Gray
            Start-Sleep -Seconds 3
            if ($elapsed.TotalSeconds -gt 20 -and -not (Test-Path -LiteralPath $result)) {
                $task = Get-ScheduledTask -CimSession $cim -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
                if ($task -and $task.State -eq 'Ready') {
                    Start-Sleep -Seconds 2
                    if (-not (Test-Path -LiteralPath $result)) {
                        $info = Get-ScheduledTaskInfo -CimSession $cim -TaskName $TaskName -TaskPath $TaskPath
                        throw "The task ended without a result (last result 0x{0:X})." -f $info.LastTaskResult
                    }
                }
            }
        }
        Write-Host ''
        $code = [int]((Get-Content -LiteralPath $result -Raw).Trim())
        $errFile = Join-Path $base 'error.txt'
        if (Test-Path -LiteralPath $errFile) { Write-Bad (Get-Content -LiteralPath $errFile -Raw).Trim() }
        Get-RemoteLogs $Computer $since $null
        return $code
    }
    finally {
        Unregister-ScheduledTask -CimSession $cim -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
        Remove-CimSession $cim -ErrorAction SilentlyContinue
        if ($Cleanup) { Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# ---------------------------------------------------------------------------
Write-Host ''
Write-Host "  Deployment Editor - Run on Remote Machine" -ForegroundColor Cyan
Write-Host "  $Label on $($Computers -join ', ')" -ForegroundColor Cyan
Write-Host ''

$script:Credential = $null
if ($AskCredential) {
    $script:Credential = Get-Credential -Message 'Account with administrator rights on the remote computer(s)'
    if (-not $script:Credential) { Write-Bad 'Cancelled.'; Read-Host '   Press Enter to close'; exit 1 }
}

$summary = @()
foreach ($Computer in $Computers) {
    Write-Host "  [$Computer]" -ForegroundColor White
    $use = $Method
    $code = $null
    $message = ''
    try {
        if ($use -eq 'Auto') {
            if ($DeployMode -eq 'Interactive') {
                $use = 'Task'
            }
            else {
                Write-Step 'Checking PowerShell Remoting (WinRM) ...'
                $use = 'Task'
                try { if (Test-WSMan -ComputerName $Computer -ErrorAction Stop) { $use = 'WinRM' } } catch { }
            }
            Write-Step ("Method: " + $(if ($use -eq 'WinRM') { 'PowerShell Remoting' } else { 'Scheduled task (SYSTEM)' }))
        }
        if ($use -eq 'WinRM' -and $DeployMode -eq 'Interactive') {
            Write-Host '   WinRM sessions have no desktop - PSADT dialogs cannot be shown to the user.' -ForegroundColor Yellow
        }
        $code = if ($use -eq 'WinRM') { Invoke-ViaWinRM $Computer } else { Invoke-ViaTask $Computer }
        $message = Get-ExitText $code
        if (Test-Success $code) { Write-Good "Exit code $code - $message" } else { Write-Bad "Exit code $code - $message" }
    }
    catch {
        $message = $_.Exception.Message
        Write-Bad $message
        $hint = Get-Hint $message
        if ($hint) { Write-Host "   $hint" -ForegroundColor Yellow }
    }
    $summary += [pscustomobject]@{
        Computer = $Computer
        Method   = $(if ($use -eq 'WinRM') { 'PowerShell Remoting' } else { 'Scheduled task' })
        ExitCode = $code
        Result   = $(if ($null -eq $code) { 'Not run: ' + $message } else { $message })
    }
    Write-Host ''
}

foreach ($d in $Drives.Values) { Remove-PSDrive -Name $d.TrimEnd(':') -Scope Script -ErrorAction SilentlyContinue }
Remove-Item -LiteralPath $Package -Force -ErrorAction SilentlyContinue

Write-Host '  Summary' -ForegroundColor Cyan
$summary | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
if ($CopyLogs -and (Test-Path -LiteralPath $LocalLogs)) {
    Write-Host "  Logs: $LocalLogs" -ForegroundColor Gray
    $answer = Read-Host '  Open the log folder? (y/N)'
    if ($answer -match '^(y|j)') { Invoke-Item -LiteralPath $LocalLogs }
}
else {
    Read-Host '  Press Enter to close' | Out-Null
}
$failed = @($summary | Where-Object { $null -eq $_.ExitCode -or -not (Test-Success ([int]$_.ExitCode)) }).Count
exit $failed
"#;

pub struct Prepared {
    pub controller: PathBuf,
    pub computers: Vec<String>,
}

/// Zip the package and write both scripts into a fresh temp folder.
pub fn prepare(project_folder: &Path, title: &str, opts: &RemoteOptions, work_root: &Path, logs_root: &Path) -> Result<Prepared> {
    let computers = computer_list(&opts.computers)?;
    let mut opts = opts.clone();
    opts.remote_dir = check_remote_dir(&opts.remote_dir)?;
    if !DEPLOY_MODES.contains(&opts.deploy_mode.as_str()) {
        bail!("Unknown deploy mode \"{}\".", opts.deploy_mode);
    }
    let dt = opts.deployment.unwrap_or(DeploymentType::Installation);
    let stamp = chrono::Local::now().format("%Y%m%d-%H%M%S").to_string();
    let key = safe_name(project_folder.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default().as_str());
    cleanup_old(work_root);
    let work = work_root.join(format!("{key}-{stamp}"));
    std::fs::create_dir_all(&work).with_context(|| format!("Creating {}", work.display()))?;
    let zip = work.join("package.zip");
    crate::intunewin::zip_folder(project_folder, &zip, &mut |_, _| {})?;
    let runner = work.join("DE-Run.ps1");
    write_ps(&runner, &runner_script(dt, &opts.deploy_mode))?;
    let logs = logs_root.join(format!("{key}-{stamp}"));
    let controller = work.join("Run-Remote.ps1");
    let script = controller_script(&ControllerInput {
        title,
        computers: &computers,
        opts: &opts,
        package_zip: &zip,
        runner: &runner,
        local_logs: &logs,
        project_key: &key,
    });
    write_ps(&controller, &script)?;
    Ok(Prepared { controller, computers })
}

/// Remove work folders of earlier runs (older than a day).
fn cleanup_old(work_root: &Path) {
    let Ok(entries) = std::fs::read_dir(work_root) else { return };
    let day = std::time::Duration::from_secs(24 * 3600);
    for e in entries.flatten() {
        let old = e.metadata().and_then(|m| m.modified()).ok().and_then(|t| t.elapsed().ok()).map(|age| age > day).unwrap_or(false);
        if old && e.path().is_dir() {
            let _ = std::fs::remove_dir_all(e.path());
        }
    }
}

pub const DEPLOY_MODES: [&str; 3] = ["Silent", "NonInteractive", "Interactive"];

/// PowerShell 5.1 reads BOM-less files as ANSI — write UTF-8 with BOM and CRLF.
fn write_ps(path: &Path, text: &str) -> Result<()> {
    let mut bytes = vec![0xEF, 0xBB, 0xBF];
    bytes.extend_from_slice(text.replace("\r\n", "\n").replace('\n', "\r\n").as_bytes());
    std::fs::write(path, bytes).with_context(|| format!("Writing {}", path.display()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn computer_names() {
        assert_eq!(computer_list("PC01, pc02\nPC-03.contoso.local;pc01").unwrap(), vec!["PC01", "pc02", "PC-03.contoso.local"]);
        assert!(computer_list("").is_err());
        assert!(computer_list("PC01; rm -rf").is_err());
        assert!(computer_list("PC'01").is_err());
        assert!(computer_list("-ComputerName").is_err());
    }

    #[test]
    fn remote_dirs() {
        assert_eq!(check_remote_dir("C:\\Windows\\Temp\\DeploymentEditor\\").unwrap(), "C:\\Windows\\Temp\\DeploymentEditor");
        assert!(check_remote_dir("C:\\").is_err());
        assert!(check_remote_dir("C:\\Temp").is_err());
        assert!(check_remote_dir("\\\\server\\share\\x").is_err());
        assert!(check_remote_dir("C:\\Temp\\$env:x").is_err());
        assert!(check_remote_dir("C:\\Temp\\..\\Windows").is_err());
    }

    #[test]
    fn scripts_are_balanced_and_complete() {
        let opts = RemoteOptions {
            computers: "PC01 PC02".into(),
            ask_credentials: true,
            deploy_mode: "Interactive".into(),
            deployment: Some(DeploymentType::Uninstall),
            ..Default::default()
        };
        let computers = computer_list(&opts.computers).unwrap();
        let s = controller_script(&ControllerInput {
            title: "7-Zip 25.01",
            computers: &computers,
            opts: &opts,
            package_zip: Path::new("C:\\Users\\it's me\\AppData\\Local\\Temp\\x\\package.zip"),
            runner: Path::new("C:\\Temp\\x\\DE-Run.ps1"),
            local_logs: Path::new("C:\\Logs"),
            project_key: "7-Zip",
        });
        assert!(s.contains("$Computers  = @('PC01', 'PC02')"));
        assert!(s.contains("$RemoteDir  = 'C:\\Windows\\Temp\\DeploymentEditor\\7-Zip'"));
        assert!(s.contains("'C:\\Users\\it''s me\\AppData"));
        assert!(s.contains("$AskCredential = $true"));
        assert!(s.contains("$Label      = 'Uninstall (Interactive)'"));
        assert_eq!(crate::validate::bracket_problem(&s), None);
        let r = runner_script(DeploymentType::Repair, "Silent");
        assert!(r.contains("-ArgumentList '-DeploymentType Repair -DeployMode Silent'"));
        assert_eq!(crate::validate::bracket_problem(&r), None);
    }

    #[test]
    fn prepare_writes_package_and_scripts() {
        let dir = std::env::temp_dir().join("de-remote-test");
        let _ = std::fs::remove_dir_all(&dir);
        let project = dir.join("My App");
        std::fs::create_dir_all(project.join("Files")).unwrap();
        std::fs::write(project.join("Invoke-AppDeployToolkit.exe"), b"exe").unwrap();
        std::fs::write(project.join("Invoke-AppDeployToolkit.db"), b"db").unwrap();
        std::fs::write(project.join("Files").join("setup.msi"), b"msi").unwrap();
        let opts = RemoteOptions { computers: "PC01".into(), ..Default::default() };
        let p = prepare(&project, "My App 1.0", &opts, &dir.join("work"), &dir.join("logs")).unwrap();
        let ps = std::fs::read(&p.controller).unwrap();
        assert!(ps.starts_with(&[0xEF, 0xBB, 0xBF]));
        let text = String::from_utf8_lossy(&ps);
        assert!(text.contains("\r\n"));
        assert!(text.contains("DeploymentEditor\\My_App'"));
        let zip = p.controller.parent().unwrap().join("package.zip");
        let mut z = zip::ZipArchive::new(std::fs::File::open(zip).unwrap()).unwrap();
        let names: Vec<String> = (0..z.len()).map(|i| z.by_index(i).unwrap().name().to_string()).collect();
        assert!(names.contains(&"Invoke-AppDeployToolkit.exe".to_string()));
        assert!(names.contains(&"Files/setup.msi".to_string()));
        assert!(!names.iter().any(|n| n.ends_with(".db")));
        assert!(prepare(&project, "x", &RemoteOptions { computers: "bad name!".into(), ..Default::default() }, &dir, &dir).is_err());
        if let Ok(out) = std::env::var("DE_REMOTE_OUT") {
            let out = std::path::Path::new(&out);
            std::fs::create_dir_all(out).unwrap();
            std::fs::copy(&p.controller, out.join("Run-Remote.ps1")).unwrap();
            std::fs::copy(p.controller.parent().unwrap().join("DE-Run.ps1"), out.join("DE-Run.ps1")).unwrap();
        }
    }
}
