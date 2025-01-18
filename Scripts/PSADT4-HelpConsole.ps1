if (Get-Module -ListAvailable -Name PSAppDeployToolkit) {
    Import-Module "PSAppDeployToolkit"

    try {
        Show-ADTHelpConsole
    } 
    catch {
        throw
    } 
    finally {
    }
}
else {
    Write-Host "The PSAppDeployToolkit is not installed. Please install it from the PowerShell Gallery first: https://www.powershellgallery.com/packages/PSAppDeployToolkit"
    Pause
}
