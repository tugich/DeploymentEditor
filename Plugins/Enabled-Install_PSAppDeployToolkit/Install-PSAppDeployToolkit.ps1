<#
    Install PSAppDeployToolkit Plugin
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath
)

try {
    Install-Module -Name PSAppDeployToolkit -Force -Verbose
    Write-Host "Successfully installed PSAppDeployToolkit module!"
    Pause
}
catch {
    throw
}
finally {
}
