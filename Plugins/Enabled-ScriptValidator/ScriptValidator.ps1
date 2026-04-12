<#
    Hello World Plugin
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath
)

function Test-PSADTScript {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        [switch]$IncludeInfo
    )

    Write-Host "`n=== Validating: $ScriptPath ===" -ForegroundColor Cyan

    # --- Step 1: Syntax Check ---
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$errors) | Out-Null

    if ($errors.Count -gt 0) {
        Write-Host "`n❌ Syntax Errors:" -ForegroundColor Red
        $errors | ForEach-Object {
            Write-Host "  [Line $($_.Extent.StartLineNumber)] $($_.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "✅ Syntax: OK" -ForegroundColor Green
    }

    # --- Step 2: Script Analyzer ---
    if (Get-Module -ListAvailable -Name PSScriptAnalyzer) {
        $severities = if ($IncludeInfo) { @('Error','Warning','Information') } else { @('Error','Warning') }
        $results = Invoke-ScriptAnalyzer -Path $ScriptPath -Severity $severities

        if ($results) {
            Write-Host "`n⚠️  Script Analyzer Issues:" -ForegroundColor Yellow
            $results | Format-Table RuleName, Severity, Line, Message -AutoSize
        } else {
            Write-Host "✅ Script Analyzer: No issues" -ForegroundColor Green
        }
    } else {
        Write-Host "⚠️  PSScriptAnalyzer not installed. Run: Install-Module PSScriptAnalyzer" -ForegroundColor Yellow
    }
}


try {
    Write-Host "Script Validator"
    Write-Host "The project path is: $ProjectPath"
	
	# Usage
    $DeploymentFile = $ProjectPath + "Invoke-AppDeployToolkit.ps1"
	Test-PSADTScript -ScriptPath $DeploymentFile -IncludeInfo  # include informational hints
	
    Pause
}
catch {
    throw
}
finally {
}
