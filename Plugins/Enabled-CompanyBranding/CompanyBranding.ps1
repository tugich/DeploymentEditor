<#
    Company Branding Plugin
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath
)

try {
    Copy-Item -Path "$PSScriptRoot\Assets" -Destination "$ProjectPath" -Recurse -Force
    Write-Host "Successfully copied branded corporate assets into the PSADT project!"
    Pause
}
catch {
    throw
}
finally {
}
