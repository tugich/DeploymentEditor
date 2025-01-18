<#
    Hello World Plugin
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath
)

try {
    Write-Host "Hello world!"
    Write-Host "The project path is: $ProjectPath"
    Pause
}
catch {
    throw
}
finally {
}
