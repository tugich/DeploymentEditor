<#
===============================================================================
 Microsoft Intune Upload Plugin
===============================================================================

.DESCRIPTION
 This script automates the packaging and upload of PowerShell App Deployment
 Toolkit (PSADT) applications as Microsoft Intune Win32 apps.

 It builds an .intunewin package on-the-fly from a specified project directory,
 extracts application metadata from the PSADT project where available, and
 uploads the application to Microsoft Intune using the IntuneWin32App PowerShell
 module and Microsoft Graph authentication (App Registration).

 The script is designed for repeatable, standardized application deployments
 and integrates seamlessly into packaging workflows and CI/CD pipelines.

 Usage by own risk.

.SYNOPSIS
 Build a PSADT-based .intunewin package in a temporary location and upload it
 as a Win32 application to Microsoft Intune.

.FEATURES
 - Automatically builds .intunewin packages using IntuneWinAppUtil
 - Supports PSADT v4 projects
 - Extracts application metadata (vendor, name, version, architecture)
 - Creates Win32 apps with registry-based detection rules
 - Applies OS and architecture requirements
 - Uploads apps using Microsoft Graph (App Registration / Client Secret)
 - Supports unattended and automated execution

.REQUIREMENTS
 - PowerShell 5.1 or PowerShell 7+
 - IntuneWin32App PowerShell module
 - Microsoft Intune tenant with Win32 app support
 - App Registration with appropriate Graph permissions
 - IntuneWinAppUtil.exe available locally

.NOTES
 Author   : Deployment Editor
 Company  : TUGI.CH
 Purpose  : Standardized, automated Intune Win32 application onboarding

===============================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath,

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "<Your Microsoft Azure Tenant ID>",

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "<Your Azure App Registration Client ID>",

    [Parameter(Mandatory = $false)]
    [SecureString]$ClientSecret,

    # IntuneWinAppUtil.exe path
    [Parameter(Mandatory = $false)]
    [string]$IntuneWinAppUtilPath = "$env:Temp\IntuneWinAppUtil.exe",

    # Setup file inside ProjectPath (PSADT default is Invoke-AppDeployToolkit.exe)
    [Parameter(Mandatory = $false)]
    [string]$SetupFile = "Invoke-AppDeployToolkit.exe",

    # Optional: If you pass a .intunewin explicitly, packaging step is skipped
    [Parameter(Mandatory = $false)]
    [string]$IntuneWinPath,

    [Parameter(Mandatory = $false)]
    [string]$AppDisplayName = "",

    [Parameter(Mandatory = $false)]
    [string]$Publisher = "Empty",

    [Parameter(Mandatory = $false)]
    [string]$Description = "Uploaded by Deployment Editor (TUGI.CH)",

    [Parameter(Mandatory = $false)]
    [string]$Version = "",

    # --- Detection (registry) ---
    [Parameter(Mandatory = $false)]
    [string]$RegistryKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Company\Deployments",

    [Parameter(Mandatory = $false)]
    [string]$DetectionValueName = "Installed",

    [Parameter(Mandatory = $false)]
    [ValidateSet("equal","notEqual","greaterThanOrEqual","greaterThan","lessThanOrEqual","lessThan")]
    [string]$DetectionOperator = "equal",

    [Parameter(Mandatory = $false)]
    [ValidatePattern("^\d+$")]
    [string]$DetectionComparisonValue = "1"
)

Set-StrictMode -Version Latest
$script:IntuneWinPath = $IntuneWinPath

# =============================================================================
# Nicer console + progress helpers
# =============================================================================
$script:StartTime = Get-Date

function Write-Banner {
    param(
        [string]$Title = "Microsoft Intune Upload Plugin",
        [string]$Sub   = "PSADT ➜ .intunewin ➜ Intune Win32"
    )
    $line = "=" * 78
    Write-Host ""
    Write-Host $line -ForegroundColor DarkGray
    Write-Host ("  {0}" -f $Title) -ForegroundColor Cyan
    Write-Host ("  {0}" -f $Sub)   -ForegroundColor Gray
    Write-Host $line -ForegroundColor DarkGray
    Write-Host ""
}

function Write-Section { param([string]$Text) Write-Host ""; Write-Host ("▶  " + $Text) -ForegroundColor Cyan }
function Write-Info    { param([string]$Text) Write-Host ("  • " + $Text) -ForegroundColor Gray }
function Write-Ok      { param([string]$Text) Write-Host ("  ✔  " + $Text) -ForegroundColor Green }
function Write-Warn    { param([string]$Text) Write-Host ("  ⚠  " + $Text) -ForegroundColor Yellow }
function Write-Fail    { param([string]$Text) Write-Host ("  ✖  " + $Text) -ForegroundColor Red }

function Set-StepProgress {
    param(
        [int]$Step,
        [int]$Total,
        [string]$Status
    )
    $pct = [int](($Step / [double]$Total) * 100)
    Write-Progress -Activity "Intune Upload" -Status $Status -PercentComplete $pct
}

function Invoke-Step {
    param(
        [Parameter(Mandatory)] [int]$Step,
        [Parameter(Mandatory)] [int]$Total,
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [scriptblock]$Action
    )
    Write-Section $Name
    Set-StepProgress -Step $Step -Total $Total -Status $Name
    try {
        & $Action
        Write-Ok $Name
    } catch {
        Write-Fail "$Name failed: $($_.Exception.Message)"
        throw
    }
}

function Finish-Progress {
    Write-Progress -Activity "Intune Upload" -Completed
    $elapsed = (Get-Date) - $script:StartTime
    Write-Host ""
    Write-Host ("Done in {0:mm\:ss}." -f $elapsed) -ForegroundColor Green
}

# =============================================================================
# Packaging helper (slightly hardened, still original intent)
# =============================================================================

function ConvertFrom-SecureStringToPlainText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [SecureString]$SecureString
    )

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Test-FilesLocked {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    # Always return an array (StrictMode-safe)
    $lockedFiles = New-Object System.Collections.Generic.List[string]

    foreach ($file in (Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue)) {
        $fullPath = $file.FullName  # capture BEFORE try/catch

        try {
            $stream = [System.IO.File]::Open(
                $fullPath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::None
            )
            $stream.Close()
        }
        catch {
            [void]$lockedFiles.Add($fullPath)
        }
    }

    return $lockedFiles.ToArray()
}


function New-IntuneWinFromProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolder,

        [Parameter(Mandatory)]
        [string]$SetupFileName,

        [Parameter(Mandatory)]
        [string]$IntuneWinAppUtilExe
    )

    $setupFullPath = Join-Path $SourceFolder $SetupFileName
    if (-not (Test-Path $setupFullPath -PathType Leaf)) {
        throw "Setup file not found: $setupFullPath. Set -SetupFile to the correct filename inside the project folder."
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outDir = Join-Path $env:TEMP ("IntuneWin\" + $stamp)
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null

    Write-Info "Preparing .intunewin in temp folder"
    Write-Info "Source: $SourceFolder"
    Write-Info "Setup : $SetupFileName"
    Write-Info "Output: $outDir"

    $argList = @(
      "-c", $SourceFolder
      "-s", $SetupFileName
      "-o", $outDir
      "-q"
    )

    # Capture logs for troubleshooting (still quiet mode)
    $stdout = Join-Path $outDir "IntuneWinAppUtil.stdout.log"
    $stderr = Join-Path $outDir "IntuneWinAppUtil.stderr.log"

    $p = Start-Process -FilePath $IntuneWinAppUtilExe -ArgumentList $argList `
        -Wait -PassThru -NoNewWindow `
        -RedirectStandardOutput $stdout `
        -RedirectStandardError  $stderr

    if ($p.ExitCode -ne 0) {
        Write-Fail "IntuneWinAppUtil.exe failed with exit code $($p.ExitCode)."
        if (Test-Path $stdout) { Write-Info "StdOut log: $stdout" }
        if (Test-Path $stderr) { Write-Info "StdErr log: $stderr" }
        throw "IntuneWinAppUtil.exe failed with exit code $($p.ExitCode)."
    }

    # DO NOT assume exact output name; take newest *.intunewin (most robust)
    $latest = Get-ChildItem -Path $outDir -Filter "*.intunewin" -File |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

    if (-not $latest) {
        Write-Fail "Packaging completed but no .intunewin was found in $outDir."
        if (Test-Path $stdout) { Write-Info "StdOut log: $stdout" }
        if (Test-Path $stderr) { Write-Info "StdErr log: $stderr" }
        throw "Packaging completed but no .intunewin was found in $outDir."
    }

    Write-Ok "Created: $($latest.FullName)"
    return $latest.FullName
}

# =============================================================================
# Check for open / locked files before packaging
# =============================================================================
Write-Banner
Write-Section "Checking for open files in project directory..."

$lockedFiles = Test-FilesLocked -Path $ProjectPath

if (@($lockedFiles).Count -gt 0) {
    Write-Warn "One or more files are currently open or locked:"
    foreach ($f in $lockedFiles) {
        Write-Host "  - $f" -ForegroundColor Red
    }
    Write-Host ""
    Write-Fail "Please close all applications/editors using these files and try again." -ForegroundColor Yellow
    Pause
    Exit 1
}

Write-Ok "No open files detected. Safe to continue packaging."

# =============================================================================
# Check for spaces in the project path
# =============================================================================
Write-Section "Checking for spaces in project path..."

if ($ProjectPath -match '\s') {
    Write-Fail "The path contains spaces. Not supported."
    Write-Host ""
    Pause
    Exit 1
}

Write-Ok "The path does NOT contain spaces."

# =============================================================================
# Main
# =============================================================================
CLS
Write-Banner
Write-Ok "Project: $ProjectPath"
Pause

$TotalSteps = 6
$step = 0

Invoke-Step -Step (++$step) -Total $TotalSteps -Name "Check prerequisites" -Action {
    # Check module IntuneWin32App (auto-install if missing)
    $moduleName = "IntuneWin32App"

    if (Get-Module -ListAvailable -Name $moduleName) {
        Write-Ok "$moduleName module is installed."
        Write-Info -Message "$moduleName module already present."
    }
    else {
        Write-Warn "$moduleName module not found. Attempting installation..."
        Write-Info -Message "$moduleName module not found. Installing..."

        try {
            # Install module (CurrentUser = no admin required)
            Install-Module `
                -Name $moduleName `
                -Scope CurrentUser `
                -Force `
                -ErrorAction Stop

            Write-Ok "$moduleName module installed successfully."
            Write-Info "$moduleName module installed successfully."

            # Import explicitly to avoid session issues
            Import-Module $moduleName -ErrorAction Stop
        }
        catch {
            Write-Fail "Failed to install $moduleName."
            throw "Required PowerShell module '$moduleName' could not be installed."
        }
    }

    # If packaging is needed, ensure IntuneWinAppUtil exists
    $needsPackaging = (-not $PSBoundParameters.ContainsKey('IntuneWinPath')) -or [string]::IsNullOrWhiteSpace($IntuneWinPath)
    if ($needsPackaging -and -not (Test-Path $IntuneWinAppUtilPath -PathType Leaf)) {
        throw "IntuneWinAppUtil.exe not found: $IntuneWinAppUtilPath"
    }

    # Ensure PSADT deployment file exists
    $DeploymentFile = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"
    if (-not (Test-Path $DeploymentFile -PathType Leaf)) {
        throw "Missing PSADT file: $DeploymentFile"
    }

    Write-Info ("Packaging required: {0}" -f $needsPackaging)
}

Invoke-Step -Step (++$step) -Total $TotalSteps -Name "Build / select .intunewin package" -Action {
    $needsPackaging = (-not $PSBoundParameters.ContainsKey('IntuneWinPath')) -or
                      [string]::IsNullOrWhiteSpace($script:IntuneWinPath)

    if ($needsPackaging) {
        $script:IntuneWinPath = New-IntuneWinFromProject `
            -SourceFolder $ProjectPath `
            -SetupFileName $SetupFile `
            -IntuneWinAppUtilExe $IntuneWinAppUtilPath
    } else {
        Write-Warn "Using provided .intunewin: $script:IntuneWinPath"
    }

    if ([string]::IsNullOrWhiteSpace($script:IntuneWinPath) -or -not (Test-Path $script:IntuneWinPath -PathType Leaf)) {
        throw "IntuneWinPath is empty or does not exist. Current value: '$script:IntuneWinPath'"
    }

    Write-Info "Package: $script:IntuneWinPath"
}

Invoke-Step -Step (++$step) -Total $TotalSteps -Name "Read PSADT metadata" -Action {
    # Read vars from Invoke-AppDeployToolkit.ps1
    $DeploymentFile = Join-Path $ProjectPath "Invoke-AppDeployToolkit.ps1"

    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $DeploymentFile, [ref]$null, [ref]$null
    )
    $hashAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left.VariablePath.UserPath -eq 'adtSession'
    }, $true)

    if (-not $hashAst) {
        throw "Could not find 'adtSession' assignment in $DeploymentFile"
    }

    $script:adtSession = Invoke-Expression $hashAst.Right.Extent.Text

    Write-Info "Vendor: $($script:adtSession.AppVendor)"
    Write-Info "Name  : $($script:adtSession.AppName)"
    Write-Info "Ver   : $($script:adtSession.AppVersion)"
    Write-Info "Arch  : $($script:adtSession.AppArch)"
    Write-Info "Lang  : $($script:adtSession.AppLang)"
}

Invoke-Step -Step (++$step) -Total $TotalSteps -Name "Connect to Microsoft Graph" -Action {

    # If you kept -ClientSecret as [SecureString], only prompt when it wasn't provided
    if (-not $PSBoundParameters.ContainsKey('ClientSecret') -or -not $ClientSecret) {
        Write-Info "Enter the Client Secret for the App Registration (input will be hidden)."
        $ClientSecret = Read-Host -Prompt "Client Secret" -AsSecureString
    }

    $plainSecret = $null
    try {
        $plainSecret = ConvertFrom-SecureStringToPlainText -SecureString $ClientSecret

        Connect-MSIntuneGraph -TenantID $TenantId -ClientID $ClientId -ClientSecret $plainSecret
        Write-Ok "Connected to Graph."
    }
    catch {
        throw "Failed connection to MS Graph. $($_.Exception.Message)"
    }
    finally {
        # Best-effort cleanup of plaintext in memory
        if ($plainSecret) { $plainSecret = $null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


Invoke-Step -Step (++$step) -Total $TotalSteps -Name "Build detection + requirement rules" -Action {
    # PSADT commands (adjust if your toolkit uses different switches)
    $script:InstallCommand   = 'Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent'
    $script:UninstallCommand = 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent'

    # Detection rule (Registry)
    $DetectionValueName = "$($script:adtSession.AppVendor) $($script:adtSession.AppName) $($script:adtSession.AppVersion) $($script:adtSession.AppArch) $($script:adtSession.AppLang)"
    $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry `
        -IntegerComparison `
        -KeyPath $RegistryKeyPath `
        -ValueName $DetectionValueName `
        -IntegerComparisonOperator $DetectionOperator `
        -IntegerComparisonValue $DetectionComparisonValue `
        -Check32BitOn64System:$false

    $script:DetectionRules = @($DetectionRule)

    # Requirement rule (OS + Arch)
    $script:RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "x64" -MinimumSupportedWindowsRelease "W11_21H2"

    Write-Info "Detection key : $RegistryKeyPath"
    Write-Info "Detection name: $DetectionValueName"
    Write-Info "Requirements  : x64, W11_21H2+"
}

Invoke-Step -Step (++$step) -Total $TotalSteps -Name "Upload Win32 app to Intune" -Action {
    # Create app
    $AppDisplayName = "$($script:adtSession.AppVendor) $($script:adtSession.AppName) $($script:adtSession.AppVersion) $($script:adtSession.AppArch)"
    $Publisher = $script:adtSession.AppVendor
    $Version = $script:adtSession.AppVersion

    $Win32Params = @{
        FilePath                = $script:IntuneWinPath
        DisplayName             = $AppDisplayName
        Description             = $Description
        Publisher               = $Publisher
        AppVersion              = $Version
        InstallCommandLine      = $script:InstallCommand
        UninstallCommandLine    = $script:UninstallCommand
        DetectionRule           = $script:DetectionRules
        RequirementRule         = $script:RequirementRule
        InstallExperience       = "system"
        RestartBehavior         = "suppress"
        AllowAvailableUninstall = $true
    }

    $App = Add-IntuneWin32App @Win32Params

    Write-Host ""
    Write-Ok "Created Win32 app"
    Write-Info "Name: $($App.displayName)"
    Write-Info "ID  : $($App.id)"
}

Finish-Progress
Pause
