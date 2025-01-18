# Specify the module name
$ModuleName = "PSAppDeployToolkit"

# Import the module (optional if it's already loaded)
Import-Module $ModuleName -ErrorAction Stop

# Get all commands from the module
$Commands = Get-Command -Module $ModuleName

# Prepare lists to store command and parameter details
$CommandList = @()
$ParameterList = @()
$CommandId = 1

foreach ($CommandItem in $Commands) {
    # Get the command description and flatten it to a single line
    $Description = ((Get-Help $CommandItem.Name -ErrorAction SilentlyContinue).Description | Out-String).Trim() -replace '\s+', ' '

    # Add command to the CommandList with an ID and description
    $CommandList += [PSCustomObject]@{
        ID          = $CommandId
        CommandName = $CommandItem.Name
        Description = $Description.Trim()
    }

    # Get the parameters of the command
    $CommandParameters = (Get-Command $CommandItem.Name).Parameters
    foreach ($ParamName in $CommandParameters.Keys) {
        $Parameter = $CommandParameters[$ParamName]

        # Extract the simplified ParameterType
        $ParameterType = $Parameter.ParameterType.Name

        # Determine if the parameter is required (1 for required, 0 for not required)
        $IsRequired = if ($Parameter.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory }) { 1 } else { 0 }

        # Get parameter description from Get-Help if available
        $ParameterHelp = (Get-Help $CommandItem.Name -ErrorAction SilentlyContinue).Parameters.Parameter | Where-Object { $_.Name -eq $ParamName }
        $ParameterDescription = ($ParameterHelp.Description | Out-String).replace("`n"," ").replace("`r"," ")

        # Fallback if no description is available
        if (-not $ParameterDescription) {
            $ParameterDescription = "No description available."
        }

        # Add each parameter to the ParameterList with the associated CommandId
        $ParameterList += [PSCustomObject]@{
            CommandID     = $CommandId
            ParameterName = $Parameter.Name
            ParameterType = $ParameterType
            IsRequired    = $IsRequired
            Description   = $ParameterDescription.Trim()
        }
    }

    # Increment CommandId for the next command
    $CommandId++
}

# Export the lists to CSV files
$CommandsCsvPath = "PSADT4_Commands.csv"
$ParametersCsvPath = "PSADT4_Parameters.csv"

$CommandList | Export-Csv -Path $CommandsCsvPath -NoTypeInformation
$ParameterList | Export-Csv -Path $ParametersCsvPath -NoTypeInformation

Write-Host "Commands exported to $CommandsCsvPath"
Write-Host "Parameters exported to $ParametersCsvPath"
