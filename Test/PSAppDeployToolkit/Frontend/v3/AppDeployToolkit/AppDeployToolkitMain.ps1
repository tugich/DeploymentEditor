<#

.SYNOPSIS
PSAppDeployToolkit - This script contains PSAppDeployToolkit v3.x API wrappers to provide backwards compatibility for Deploy-Application.ps1 scripts against PSAppDeployToolkit v4.

.DESCRIPTION
The script can be called directly to dot-source the toolkit functions for testing, but it is usually called by the Deploy-Application.ps1 script.

The script can usually be updated to the latest version without impacting your per-application Deploy-Application scripts. Please check release notes before upgrading.

PSAppDeployToolkit is licensed under the GNU LGPLv3 License - © 2025 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None. This script does not generate any output.

.LINK
https://psappdeploytoolkit.com

#>

#---------------------------------------------------------------------------
#
# MARK: Initialization code
#
#---------------------------------------------------------------------------

[CmdletBinding()]
param
(
)

# Remove all functions defined in this script from the function provider.
Remove-Item -LiteralPath ($adtWrapperFuncs = $MyInvocation.MyCommand.ScriptBlock.Ast.EndBlock.Statements | & { process { if ($_ -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return "Microsoft.PowerShell.Core\Function::$($_.Name)" } } }) -Force -ErrorAction Ignore


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Write-ADTLogEntry
#
#---------------------------------------------------------------------------

function Write-Log
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = "Apparently 'Write-Log' was a shipped cmdlet in PowerShell Core 6.1.x. We can't rename this wrapper so we must suppress.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [Alias('Text')]
        [System.String[]]$Message,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateRange(0, 3)]
        [System.Int16]$Severity,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Source = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ScriptSection = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateSet('CMTrace', 'Legacy')]
        [System.String]$LogType = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogFileDirectory = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 6)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogFileName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 7)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$AppendToLogFile,

        [Parameter(Mandatory = $false, Position = 8)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$MaxLogHistory,

        [Parameter(Mandatory = $false, Position = 9)]
        [ValidateNotNullOrEmpty()]
        [System.Decimal]$MaxLogFileSizeMB,

        [Parameter(Mandatory = $false, Position = 10)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true,

        [Parameter(Mandatory = $false, Position = 11)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$WriteHost,

        [Parameter(Mandatory = $false, Position = 12)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false, Position = 13)]
        [System.Management.Automation.SwitchParameter]$DebugMessage,

        [Parameter(Mandatory = $false, Position = 14)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$LogDebugMessage
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce overall deprecation.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Write-ADTLogEntry]. Please migrate your scripts to use the new function." -Severity 2 -Source $MyInvocation.MyCommand.Name -DebugMessage:$noDepWarnings

        # Announce dead parameters.
        $null = ('AppendToLogFile', 'MaxLogHistory', 'MaxLogFileSizeMB', 'WriteHost', 'LogDebugMessage').ForEach({
                if ($PSBoundParameters.ContainsKey($_))
                {
                    Write-ADTLogEntry -Message "The parameter '-$_' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
                    $PSBoundParameters.Remove($_)
                }
            })

        # There should never be a time where we can't log.
        if ($PSBoundParameters.ContainsKey('ContinueOnError'))
        {
            $null = $PSBoundParameters.Remove('ContinueOnError')
        }

        # Set up collector for piped in messages.
        $messages = [System.Collections.Generic.List[System.String]]::new()
    }

    process
    {
        # Add all non-null messages to the collector.
        $Message | & {
            process
            {
                if (![System.String]::IsNullOrWhiteSpace($_))
                {
                    $messages.Add($_)
                }
            }
        }
    }

    end
    {
        # Process provided messages if we have any.
        if ($messages.Count)
        {
            try
            {
                $PSBoundParameters.Message = $messages
                Write-ADTLogEntry @PSBoundParameters
            }
            catch
            {
                if (!$ContinueOnError)
                {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            }
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Close-ADTSession
#
#---------------------------------------------------------------------------

function Exit-Script
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$ExitCode
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Close-ADTSession]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Close-ADTSession @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Invoke-ADTAllUsersRegistryAction
#
#---------------------------------------------------------------------------

function Invoke-HKCURegistrySettingsForAllUsers
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ if ($_ -match '\$UserProfile\.SID') { Write-ADTLogEntry -Message "The base function [Invoke-ADTAllUsersRegistryAction] no longer supports the use of [`$UserProfile]. Please use [`$_] or [`$PSItem] instead." -Severity 2 }; ![System.String]::IsNullOrWhiteSpace($_) })]
        [Alias('RegistrySettings')]
        [System.Management.Automation.ScriptBlock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.UserProfile[]]$UserProfiles
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Invoke-ADTAllUsersRegistryAction]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    $PSBoundParameters.ScriptBlock = { New-Variable -Name UserProfile -Value $_ -Force }, $PSBoundParameters.ScriptBlock
    try
    {
        Invoke-ADTAllUsersRegistryAction @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Replacement for Get-HardwarePlatform
#
#---------------------------------------------------------------------------

function Get-HardwarePlatform
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [`$envHardwareType]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        return $envHardwareType
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTFreeDiskSpace
#
#---------------------------------------------------------------------------

function Get-FreeDiskSpace
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Drive = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTFreeDiskSpace]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Get-ADTFreeDiskSpace @PSBoundParameters
    }
    catch
    {
        Write-ADTLogEntry -Message "Failed to retrieve free disk space for drive [$Drive].`n$(Resolve-ADTErrorRecord -ErrorRecord $_)" -Severity 3
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Remove-ADTInvalidFileNameChars
#
#---------------------------------------------------------------------------

function Remove-InvalidFileNameChars
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [System.String]$Name
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce deprecation of function and set up accumulator for all piped in names.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Remove-ADTInvalidFileNameChars]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
        $names = [System.Collections.Generic.List[System.String]]::new()
    }

    process
    {
        # Add all non-null names to the collector.
        if (![System.String]::IsNullOrWhiteSpace($Name))
        {
            $names.Add($Name)
        }
    }

    end
    {
        # Process provided names if we have any.
        if ($names.Count)
        {
            try
            {
                $names | Remove-ADTInvalidFileNameChars
            }
            catch
            {
                $PSCmdlet.ThrowTerminatingError($_)
            }
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTApplication
#
#---------------------------------------------------------------------------

function Get-InstalledApplication
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Name', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ProductCode', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'IncludeUpdatesAndHotfixes', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Name,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ProductCode = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Exact,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$WildCard,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$RegEx,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeUpdatesAndHotfixes
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTApplication]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    $gaiaParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation -Exclude Exact, WildCard, RegEx

    if ($Exact)
    {
        $gaiaParams.NameMatch = 'Exact'
    }
    elseif ($WildCard)
    {
        $gaiaParams.NameMatch = 'WildCard'
    }
    elseif ($RegEx)
    {
        $gaiaParams.NameMatch = 'RegEx'
    }

    # Invoke execution.
    try
    {
        Get-ADTApplication @gaiaParams
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Uninstall-ADTApplication
#
#---------------------------------------------------------------------------

function Remove-MSIApplications
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Name', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidAssignmentToAutomaticVariable', '', Justification = '$_ is intentionally overwritten in this function to expand the input array.')]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ArgumentList', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'AdditionalArgumentList', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'IncludeUpdatesAndHotfixes', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'LoggingOptions', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'LogFileName', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'PassThru', Justification = "This parameter is passed to an underlying function via `$PSBoundParameters, therefore this warning is benign.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Exact', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'WildCard', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Exact,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$WildCard,

        [Parameter(Mandatory = $false)]
        [Alias('Arguments', 'Parameters')]
        [ValidateNotNullOrEmpty()]
        [System.String]$ArgumentList = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Alias('AddParameters')]
        [System.String]$AdditionalArgumentList = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Array]$FilterApplication,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Array]$ExcludeFromUninstall,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeUpdatesAndHotfixes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LoggingOptions = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [Alias('LogName')]
        [System.String]$LogFileName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Uninstall-ADTApplication]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Build out hashtable for splatting.
    $uaaParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation -Exclude Exact, WildCard, FilterApplication, ExcludeFromUninstall, ContinueOnError
    $uaaParams.ApplicationType = 'MSI'
    if (!$ContinueOnError)
    {
        $uaaParams.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    # Build out filterscript based on provided input.
    $filterArray = $(
        $filterApplication | & {
            process
            {
                if ($null -ne $_)
                {
                    if ($_.Count -eq 1 -and $_[0].Count -eq 3) { $_ = $_[0] } # Handle the case where input is of the form @(, @('Prop', 'Value', 'Exact'), @('Prop', 'Value', 'Exact'))
                    if ($_[2] -eq 'RegEx')
                    {
                        "`$_.$($_[0]) -match '$($_[1] -replace "'","''")'"
                    }
                    elseif ($_[2] -eq 'Contains')
                    {
                        "`$_.$($_[0]) -match '$([System.Text.RegularExpressions.Regex]::Escape(($_[1] -replace "'","''")))'"
                    }
                    elseif ($_[2] -eq 'WildCard')
                    {
                        "`$_.$($_[0]) -like '$($_[1] -replace "'","''")'"
                    }
                    elseif ($_[2] -eq 'Exact')
                    {
                        if ($_[1] -is [System.Boolean])
                        {
                            "`$_.$($_[0]) -eq `$$($_[1].ToString().ToLowerInvariant())"
                        }
                        else
                        {
                            "`$_.$($_[0]) -eq '$($_[1] -replace "'","''")'"
                        }
                    }
                }
            }
        }
        $excludeFromUninstall | & {
            process
            {
                if ($null -ne $_)
                {
                    if ($_.Count -eq 1 -and $_[0].Count -eq 3) { $_ = $_[0] } # Handle the case where input is of the form @(, @('Prop', 'Value', 'Exact'), @('Prop', 'Value', 'Exact'))
                    if ($_[2] -eq 'RegEx')
                    {
                        "`$_.$($_[0]) -notmatch '$($_[1] -replace "'","''")'"
                    }
                    elseif ($_[2] -eq 'Contains')
                    {
                        "`$_.$($_[0]) -notmatch '$([System.Text.RegularExpressions.Regex]::Escape(($_[1] -replace "'","''")))'"
                    }
                    elseif ($_[2] -eq 'WildCard')
                    {
                        "`$_.$($_[0]) -notlike '$($_[1] -replace "'","''")'"
                    }
                    elseif ($_[2] -eq 'Exact')
                    {
                        if ($_[1] -is [System.Boolean])
                        {
                            "`$_.$($_[0]) -ne `$$($_[1].ToString().ToLowerInvariant())"
                        }
                        else
                        {
                            "`$_.$($_[0]) -ne '$($_[1] -replace "'","''")'"
                        }
                    }
                }
            }
        }
    )

    $filterScript = [System.String]::Join(' -and ', $filterArray)

    if ($filterScript)
    {
        $uaaParams.filterScript = [System.Management.Automation.ScriptBlock]::Create($filterScript)
    }

    # Invoke execution.
    try
    {
        Uninstall-ADTApplication @uaaParams
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}

#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTFileVersion
#
#---------------------------------------------------------------------------

function Get-FileVersion
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$File,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ProductVersion,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTFileVersion]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Get-ADTFileVersion @PSBoundParameters
    }
    catch
    {
        Write-ADTLogEntry -Message "Failed to get version info.`n$(Resolve-ADTErrorRecord -ErrorRecord $_)" -Severity 3
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTUserProfiles
#
#---------------------------------------------------------------------------

function Get-UserProfiles
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ExcludeNTAccount,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExcludeSystemProfiles = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExcludeServiceProfiles = $true,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeDefaultUser
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Translate parameters.
    $null = ('SystemProfiles', 'ServiceProfiles').Where({ $PSBoundParameters.ContainsKey("Exclude$_") }).ForEach({
            if (!$PSBoundParameters."Exclude$_")
            {
                $PSBoundParameters.Add("Include$_", [System.Management.Automation.SwitchParameter]$true)
            }
            $PSBoundParameters.Remove("Exclude$_")
        })

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTUserProfiles]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Get-ADTUserProfiles @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Update-ADTDesktop
#
#---------------------------------------------------------------------------

function Update-Desktop
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Update-ADTDesktop]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Update-ADTDesktop
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Update-ADTEnvironmentPsProvider
#
#---------------------------------------------------------------------------

function Update-SessionEnvironmentVariables
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [System.Management.Automation.SwitchParameter]$LoadLoggedOnUserEnvironmentVariables,

        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Update-ADTEnvironmentPsProvider]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Update-ADTEnvironmentPsProvider -LoadLoggedOnUserEnvironmentVariables:$LoadLoggedOnUserEnvironmentVariables
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Copy-ADTFile
#
#---------------------------------------------------------------------------

function Copy-File
{
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Path,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Destination,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Flatten,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueFileCopyOnError = $false,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$UseRobocopy,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$RobocopyParams = '/NJH /NJS /NS /NC /NP /NDL /FP /IS /IT /IM /XX /MT:4 /R:1 /W:1',

        [Parameter(Mandatory = $false)]
        [System.String]$RobocopyAdditionalParams
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Copy-ADTFile]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    if (!$UseRobocopy)
    {
        if ($PSBoundParameters.ContainsKey('RobocopyParams'))
        {
            $null = $PSBoundParameters.Remove('RobocopyParams')
        }
        if ($PSBoundParameters.ContainsKey('RobocopyAdditionalParams'))
        {
            $null = $PSBoundParameters.Remove('RobocopyAdditionalParams')
        }
    }
    if ($PSBoundParameters.ContainsKey('UseRobocopy'))
    {
        $null = $PSBoundParameters.Add('FileCopyMode', ('Native', 'Robocopy')[$PSBoundParameters.UseRobocopy])
        $null = $PSBoundParameters.Remove('UseRobocopy')
    }
    try
    {
        Copy-ADTFile @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Remove-ADTFile
#
#---------------------------------------------------------------------------

function Remove-File
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Path,

        [Parameter(Mandatory = $true, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$LiteralPath,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Remove-ADTFile]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Remove-ADTFile @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Copy-ADTFileToUserProfiles
#
#---------------------------------------------------------------------------

function Copy-FileToUserProfiles
{
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
        [System.String[]]$Path,

        [Parameter(Mandatory = $false, Position = 2)]
        [System.String]$Destination = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Flatten,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$UseRobocopy,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$RobocopyAdditionalParams = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ExcludeNTAccount,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExcludeSystemProfiles = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExcludeServiceProfiles = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$ExcludeDefaultUser,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueFileCopyOnError
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Copy-ADTFileToUserProfiles]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
        $null = ('SystemProfiles', 'ServiceProfiles').Where({ $PSBoundParameters.ContainsKey("Exclude$_") }).ForEach({
                if (!$PSBoundParameters."Exclude$_")
                {
                    $PSBoundParameters.Add("Include$_", [System.Management.Automation.SwitchParameter]$true)
                }
                $PSBoundParameters.Remove("Exclude$_")
            })
        if ($PSBoundParameters.ContainsKey('UseRobocopy'))
        {
            $PSBoundParameters.Add('FileCopyMode', ('Native', 'Robocopy')[$PSBoundParameters.UseRobocopy])
            $null = $PSBoundParameters.Remove('UseRobocopy')
        }
        if ($PSBoundParameters.ContainsKey('ContinueOnError'))
        {
            $null = $PSBoundParameters.Remove('ContinueOnError')
        }
        if (!$ContinueOnError)
        {
            $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
        }

        # Set up collector for piped in path objects.
        $srcPaths = [System.Collections.Generic.List[System.String]]::new()
    }

    process
    {
        # Add all non-null strings to the collector.
        $Path | & {
            process
            {
                if (![System.String]::IsNullOrWhiteSpace($_))
                {
                    $srcPaths.Add($_)
                }
            }
        }
    }

    end
    {
        # Process provided paths if we have any.
        if ($srcPaths.Count)
        {
            try
            {
                $PSBoundParameters.Path = $srcPaths
                Copy-ADTFileToUserProfiles @PSBoundParameters
            }
            catch
            {
                if (!$ContinueOnError)
                {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            }
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Show-ADTInstallationPrompt
#
#---------------------------------------------------------------------------

function Show-InstallationPrompt
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Left', 'Center', 'Right')]
        [System.String]$MessageAlignment = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonRightText = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonLeftText = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonMiddleText = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Application', 'Asterisk', 'Error', 'Exclamation', 'Hand', 'Information', 'None', 'Question', 'Shield', 'Warning', 'WinLogo')]
        [System.String]$Icon = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PersistPrompt,

        [Parameter(Mandatory = $false)]
        [System.Boolean]$MinimizeWindows = $false,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.UInt32]]$Timeout,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExitOnTimeout,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$TopMost
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Show-ADTInstallationPrompt]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Tune up parameters. A lot has changed.
    if ($PSBoundParameters.ContainsKey('Icon') -and ($PSBoundParameters.Icon -eq 'None'))
    {
        $null = $PSBoundParameters.Remove('Icon')
    }
    if ($PSBoundParameters.ContainsKey('ExitOnTimeout'))
    {
        $PSBoundParameters.Add('NoExitOnTimeout', !$PSBoundParameters.ExitOnTimeout)
        $null = $PSBoundParameters.Remove('ExitOnTimeout')
    }
    if ($PSBoundParameters.ContainsKey('TopMost'))
    {
        $PSBoundParameters.Add('NotTopMost', !$PSBoundParameters.TopMost)
        $null = $PSBoundParameters.Remove('TopMost')
    }

    # Invoke function with amended parameters.
    try
    {
        Show-ADTInstallationPrompt @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Show-ADTInstallationProgress
#
#---------------------------------------------------------------------------

function Show-InstallationProgress
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$StatusMessage = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Default', 'TopLeft', 'Top', 'TopRight', 'TopCenter', 'BottomLeft', 'Bottom', 'BottomRight')]
        [System.String]$WindowLocation = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$TopMost = $true,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Quiet,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoRelocation
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Show-ADTInstallationProgress]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('TopMost'))
    {
        $PSBoundParameters.Add('NotTopMost', !$PSBoundParameters.TopMost)
        $null = $PSBoundParameters.Remove('TopMost')
    }
    if ($PSBoundParameters.ContainsKey('Quiet'))
    {
        $PSBoundParameters.Add('InformationAction', [System.Management.Automation.ActionPreference]::SilentlyContinue)
        $null = $PSBoundParameters.Remove('Quiet')
    }
    try
    {
        Show-ADTInstallationProgress @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Show-ADTDialogBox
#
#---------------------------------------------------------------------------

function Show-DialogBox
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Enter a message for the dialog box.')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Text,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('OK', 'OKCancel', 'AbortRetryIgnore', 'YesNoCancel', 'YesNo', 'RetryCancel', 'CancelTryAgainContinue')]
        [System.String]$Buttons = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('First', 'Second', 'Third')]
        [System.String]$DefaultButton = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Exclamation', 'Information', 'None', 'Stop', 'Question')]
        [System.String]$Icon = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Timeout = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$TopMost
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Show-ADTDialogBox]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('TopMost'))
    {
        $PSBoundParameters.Add('NotTopMost', !$PSBoundParameters.TopMost)
        $null = $PSBoundParameters.Remove('TopMost')
    }
    try
    {
        Show-ADTDialogBox @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Show-ADTInstallationWelcome
#
#---------------------------------------------------------------------------

function Show-InstallationWelcome
{
    [CmdletBinding(DefaultParameterSetName = 'None')]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$CloseApps = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Silent,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$CloseAppsCountdown,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$ForceCloseAppsCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PromptToSave,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PersistPrompt,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$BlockExecution,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$AllowDefer,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$AllowDeferCloseApps,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$DeferTimes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$DeferDays,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DeferDeadline = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(ParameterSetName = 'CheckDiskSpaceParameterSet', Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CheckDiskSpace,

        [Parameter(ParameterSetName = 'CheckDiskSpaceParameterSet', Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$RequiredDiskSpace,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$MinimizeWindows = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$TopMost = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$ForceCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CustomText
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Show-ADTInstallationWelcome]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Tune up parameters. A lot has changed.
    if ($PSBoundParameters.ContainsKey('CloseApps'))
    {
        $PSBoundParameters.CloseProcesses = $CloseApps.Split(',') | & {
            process
            {
                $name, $description = $_.Split('=')
                if ($description)
                {
                    return [PSADT.ProcessManagement.ProcessDefinition]::new($name, $description)
                }
                else
                {
                    return [PSADT.ProcessManagement.ProcessDefinition]::new($name)
                }
            }
        }
        $null = $PSBoundParameters.Remove('CloseApps')
    }
    $null = ('{0}Countdown', 'Force{0}Countdown', 'AllowDefer{0}').ForEach({
            if ($PSBoundParameters.ContainsKey(($oldParam = [System.String]::Format($_, 'CloseApps'))))
            {
                $PSBoundParameters.Add([System.String]::Format($_, 'CloseProcesses'), $PSBoundParameters.$oldParam)
                $PSBoundParameters.Remove($oldParam)
            }
        })
    if ($PSBoundParameters.ContainsKey('TopMost'))
    {
        $PSBoundParameters.Add('NotTopMost', !$PSBoundParameters.TopMost)
        $null = $PSBoundParameters.Remove('TopMost')
    }
    if ($MinimizeWindows)
    {
        $PSBoundParameters.Add('MinimizeWindows', $MinimizeWindows)
    }

    # Invoke function with amended parameters.
    try
    {
        Show-ADTInstallationWelcome @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTWindowTitle
#
#---------------------------------------------------------------------------

function Get-WindowTitle
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'SearchWinTitle')]
        [AllowEmptyString()]
        [System.String]$WindowTitle,

        [Parameter(Mandatory = $true, ParameterSetName = 'GetAllWinTitles')]
        [System.Management.Automation.SwitchParameter]$GetAllWindowTitles,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DisableFunctionLogging
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTWindowTitle]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('DisableFunctionLogging'))
    {
        $PSBoundParameters.Add('InformationAction', [System.Management.Automation.ActionPreference]::SilentlyContinue)
        $null = $PSBoundParameters.Remove('DisableFunctionLogging')
    }
    try
    {
        Get-ADTWindowTitle @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Show-ADTInstallationRestartPrompt
#
#---------------------------------------------------------------------------

function Show-InstallationRestartPrompt
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$CountdownSeconds,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$CountdownNoHideSeconds,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$NoSilentRestart = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$SilentCountdownSeconds,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$TopMost = $true,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoCountdown
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Show-ADTInstallationRestartPrompt]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('NoSilentRestart'))
    {
        $PSBoundParameters.Add('SilentRestart', !$PSBoundParameters.NoSilentRestart)
        $null = $PSBoundParameters.Remove('NoSilentRestart')
    }
    if ($PSBoundParameters.ContainsKey('TopMost'))
    {
        $PSBoundParameters.Add('NotTopMost', !$PSBoundParameters.TopMost)
        $null = $PSBoundParameters.Remove('TopMost')
    }
    try
    {
        Show-ADTInstallationRestartPrompt @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Show-ADTBalloonTip
#
#---------------------------------------------------------------------------

function Show-BalloonTip
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]$BalloonTipText,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]$BalloonTipTitle = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateSet('Error', 'Info', 'None', 'Warning')]
        [System.Windows.Forms.ToolTipIcon]$BalloonTipIcon,

        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$BalloonTipTime,

        [Parameter(Mandatory = $false, Position = 4)]
        [System.Management.Automation.SwitchParameter]$NoWait
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Show-ADTBalloonTip]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($NoWait)
    {
        Write-ADTLogEntry -Message "The parameter '-NoWait' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
        $null = $PSBoundParameters.Remove('NoWait')
    }
    try
    {
        Show-ADTBalloonTip @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Copy-ADTContentToCache
#
#---------------------------------------------------------------------------

function Copy-ContentToCache
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'The path to the software cache folder')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Copy-ADTContentToCache]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Copy-ADTContentToCache @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Remove-ADTContentFromCache
#
#---------------------------------------------------------------------------

function Remove-ContentFromCache
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'The path to the software cache folder')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Remove-ADTContentFromCache]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Remove-ADTContentFromCache @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTNetworkConnection
#
#---------------------------------------------------------------------------

function Test-NetworkConnection
{
    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-ADTNetworkConnection]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Test-ADTNetworkConnection
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTLoggedOnUser
#
#---------------------------------------------------------------------------

function Get-LoggedOnUser
{
    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTLoggedOnUser]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Get-ADTLoggedOnUser
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTIniValue
#
#---------------------------------------------------------------------------

function Get-IniValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!(Test-Path -LiteralPath $_ -PathType Leaf))
                {
                    $PSCmdlet.ThrowTerminatingError((New-ADTValidateScriptErrorRecord -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Section,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTIniValue]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Get-ADTIniValue @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTIniValue
#
#---------------------------------------------------------------------------

function Set-IniValue
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!(Test-Path -LiteralPath $_ -PathType Leaf))
                {
                    $PSCmdlet.ThrowTerminatingError((New-ADTValidateScriptErrorRecord -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Section,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [System.Object]$Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTIniValue]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Set-ADTIniValue @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around New-ADTFolder
#
#---------------------------------------------------------------------------

function New-Folder
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [New-ADTFolder]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        New-ADTFolder @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTPowerPoint
#
#---------------------------------------------------------------------------

function Test-PowerPoint
{
    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-PowerPoint]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Test-ADTPowerPoint
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Update-ADTGroupPolicy
#
#---------------------------------------------------------------------------

function Update-GroupPolicy
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Update-ADTGroupPolicy]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Update-ADTGroupPolicy @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTUniversalDate
#
#---------------------------------------------------------------------------

function Get-UniversalDate
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DateTime = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $false
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTUniversalDate]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Get-ADTUniversalDate @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTServiceExists
#
#---------------------------------------------------------------------------

function Test-ServiceExists
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ComputerName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-ADTServiceExists]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($ComputerName)
    {
        Write-ADTLogEntry -Message "The parameter '-ComputerName' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
        $null = $PSBoundParameters.Remove('ComputerName')
    }
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Test-ADTServiceExists @PSBoundParameters -UseCIM
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Disable-ADTTerminalServerInstallMode
#
#---------------------------------------------------------------------------

function Disable-TerminalServerInstallMode
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Disable-ADTTerminalServerInstallMode]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Disable-ADTTerminalServerInstallMode @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Disable-ADTTerminalServerInstallMode
#
#---------------------------------------------------------------------------

function Enable-TerminalServerInstallMode
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Enable-ADTTerminalServerInstallMode]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Enable-ADTTerminalServerInstallMode @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Add-ADTEdgeExtension and Remove-ADTEdgeExtension
#
#---------------------------------------------------------------------------

function Configure-EdgeExtension
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = "This compatibility wrapper function cannot have its name changed for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Add')]
        [System.Management.Automation.SwitchParameter]$Add,

        [Parameter(Mandatory = $true, ParameterSetName = 'Remove')]
        [System.Management.Automation.SwitchParameter]$Remove,

        [Parameter(Mandatory = $true, ParameterSetName = 'Add')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Remove')]
        [ValidateNotNullOrEmpty()]
        [System.String]$ExtensionID,

        [Parameter(Mandatory = $true, ParameterSetName = 'Add')]
        [ValidateSet('blocked', 'allowed', 'removed', 'force_installed', 'normal_installed')]
        [System.String]$InstallationMode,

        [Parameter(Mandatory = $true, ParameterSetName = 'Add')]
        [ValidateNotNullOrEmpty()]
        [System.String]$UpdateUrl,

        [Parameter(Mandatory = $false, ParameterSetName = 'Add')]
        [ValidateNotNullOrEmpty()]
        [System.String]$MinimumVersionRequired
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [$($PSCmdlet.ParameterSetName)-ADTEdgeExtension]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    $null = $PSBoundParameters.Remove($PSCmdlet.ParameterSetName)
    try
    {
        & "$($PSCmdlet.ParameterSetName)-ADTEdgeExtension" @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Resolve-ADTErrorRecord
#
#---------------------------------------------------------------------------

function Resolve-Error
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidDefaultValueSwitchParameter', '', Justification = "This compatibility layer has several switches defaulting to True out of necessity for supporting PSAppDeployToolit 3.x Deploy-Application.ps1 scripts.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [System.Array]$ErrorRecord,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Property,

        [Parameter(Mandatory = $false, Position = 2)]
        [System.Management.Automation.SwitchParameter]$GetErrorRecord = $true,

        [Parameter(Mandatory = $false, Position = 3)]
        [System.Management.Automation.SwitchParameter]$GetErrorInvocation = $true,

        [Parameter(Mandatory = $false, Position = 4)]
        [System.Management.Automation.SwitchParameter]$GetErrorException = $true,

        [Parameter(Mandatory = $false, Position = 5)]
        [System.Management.Automation.SwitchParameter]$GetErrorInnerException = $true
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce overall deprecation and translate bad switches before executing.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Resolve-ADTErrorRecord]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
        $null = ('ErrorRecord', 'ErrorInvocation', 'ErrorException', 'ErrorInnerException').Where({ $PSBoundParameters.ContainsKey($_) }).ForEach({
                $PSBoundParameters.Add("Exclude$_", !$PSBoundParameters."Get$_")
                $PSBoundParameters.Remove("Get$_")
            })

        # Set up collector for piped in ErrorRecord objects.
        $errRecords = [System.Collections.Generic.List[System.Management.Automation.ErrorRecord]]::new()
    }

    process
    {
        # Process piped input and collect ErrorRecord objects.
        $ErrorRecord | & {
            process
            {
                if ($_ -is [System.Management.Automation.ErrorRecord])
                {
                    $errRecords.Add($_)
                }
            }
        }
    }

    end
    {
        # Process the collected ErrorRecord objects.
        try
        {
            # If we've collected no ErrorRecord objects, choose the latest error that occurred.
            if (!$errRecords.Count)
            {
                if (($errRecord = Get-Variable -Name PSItem -Scope 1 -ValueOnly -ErrorAction Ignore) -and ($errRecord -is [System.Management.Automation.ErrorRecord]))
                {
                    $errRecord | Resolve-ADTErrorRecord @PSBoundParameters
                }
                elseif ($Global:Error.Count)
                {
                    $Global:Error.Where({ $_ -is [System.Management.Automation.ErrorRecord] }, 'First', 1) | Resolve-ADTErrorRecord @PSBoundParameters
                }
            }
            else
            {
                if ($PSBoundParameters.ContainsKey('ErrorRecord'))
                {
                    $null = $PSBoundParameters.Remove('ErrorRecord')
                }
                $errRecords | Resolve-ADTErrorRecord @PSBoundParameters
            }
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTServiceStartMode
#
#---------------------------------------------------------------------------

function Get-ServiceStartMode
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String]$Service,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ComputerName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTServiceStartMode]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($ComputerName)
    {
        Write-ADTLogEntry -Message "The parameter '-ComputerName' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
        $null = $PSBoundParameters.Remove('ComputerName')
    }
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    try
    {
        Get-ADTServiceStartMode @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTServiceStartMode
#
#---------------------------------------------------------------------------

function Set-ServiceStartMode
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String]$Service,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$StartMode = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTServiceStartMode]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    try
    {
        Set-ADTServiceStartMode @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Start-ADTProcess
#
#---------------------------------------------------------------------------

function Execute-Process
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = "This compatibility wrapper function cannot have its name changed for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [Alias('Path')]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [Alias('Arguments', 'Parameters')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ArgumentList,

        [Parameter(Mandatory = $false)]
        [Alias('SecureParameters')]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Normal', 'Hidden', 'Maximized', 'Minimized')]
        [System.Diagnostics.ProcessWindowStyle]$WindowStyle,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CreateNoWindow,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$WorkingDirectory,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$WaitForMsiExec,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$MsiExecWaitTime = (Get-ADTConfig).MSI.MutexWaitTime,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IgnoreExitCodes = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Idle', 'Normal', 'High', 'AboveNormal', 'BelowNormal', 'RealTime')]
        [System.Diagnostics.ProcessPriorityClass]$PriorityClass,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExitOnProcessFailure = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$UseShellExecute,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $false
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce deprecation of this function.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Start-ADTProcess]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Convert out changed parameters.
    if ($PSBoundParameters.ContainsKey('MsiExecWaitTime'))
    {
        $PSBoundParameters.MsiExecWaitTime = [System.TimeSpan]::FromSeconds($MsiExecWaitTime)
    }
    if ($PSBoundParameters.ContainsKey('IgnoreExitCodes'))
    {
        $PSBoundParameters.IgnoreExitCodes = $IgnoreExitCodes.Split(',')
    }
    if ($PSBoundParameters.ContainsKey('ContinueOnError') -or $PSBoundParameters.ContainsKey('ExitOnProcessFailure'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
        $null = $PSBoundParameters.Remove('ExitOnProcessFailure')
        $PSBoundParameters.ErrorAction = ([System.Management.Automation.ActionPreference]::Stop, [System.Management.Automation.ActionPreference]::SilentlyContinue)[$ContinueOnError -or !$ExitOnProcessFailure]
    }

    # Invoke function with amended parameters.
    try
    {
        Start-ADTProcess @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Start-ADTMsiProcess
#
#---------------------------------------------------------------------------

function Execute-MSI
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = "This compatibility wrapper function cannot have its name changed for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Install', 'Uninstall', 'Patch', 'Repair', 'ActiveSetup')]
        [System.String]$Action = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $true, HelpMessage = 'Please enter either the path to the MSI/MSP file or the ProductCode')]
        [ValidateScript({ ($_ -match (Get-ADTEnvironmentTable).MSIProductCodeRegExPattern) -or ('.msi', '.msp' -contains [System.IO.Path]::GetExtension($_)) })]
        [Alias('Path')]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Transform = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Alias('Arguments', 'Parameters')]
        [System.String]$ArgumentList = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Alias('AddParameters')]
        [System.String]$AdditionalArgumentList = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [Alias('SecureParameters')]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Patch = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LoggingOptions = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Alias('LogName')]
        [System.String]$LogFileName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipMSIAlreadyInstalledCheck,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeUpdatesAndHotfixes,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IgnoreExitCodes = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Idle', 'Normal', 'High', 'AboveNormal', 'BelowNormal', 'RealTime')]
        [Diagnostics.ProcessPriorityClass]$PriorityClass,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExitOnProcessFailure = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$RepairFromSource,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $false
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce deprecation of this function.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Start-ADTMsiProcess]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Convert out changed parameters.
    if ($FilePath -match (Get-ADTEnvironmentTable).MSIProductCodeRegExPattern)
    {
        $PSBoundParameters.ProductCode = [System.Guid]::new($FilePath)
        $null = $PSBoundParameters.Remove('FilePath')
    }
    if ($PSBoundParameters.ContainsKey('Transform'))
    {
        $PSBoundParameters.Transforms = $Transform.Split(';')
        $null = $PSBoundParameters.Remove('Transform')
    }
    if ($PSBoundParameters.ContainsKey('IgnoreExitCodes'))
    {
        $PSBoundParameters.IgnoreExitCodes = $IgnoreExitCodes.Split(',')
    }
    if ($PSBoundParameters.ContainsKey('ContinueOnError') -or $PSBoundParameters.ContainsKey('ExitOnProcessFailure'))
    {
        $PSBoundParameters.ErrorAction = ([System.Management.Automation.ActionPreference]::Stop, [System.Management.Automation.ActionPreference]::SilentlyContinue)[$ContinueOnError -or !$ExitOnProcessFailure]
        $null = $PSBoundParameters.Remove('ContinueOnError')
        $null = $PSBoundParameters.Remove('ExitOnProcessFailure')
    }

    # Invoke function with amended parameters.
    try
    {
        Start-ADTMsiProcess @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Start-ADTMspProcess
#
#---------------------------------------------------------------------------

function Execute-MSP
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = "This compatibility wrapper function cannot have its name changed for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, HelpMessage = 'Please enter the path to the MSP file')]
        [ValidateScript({ ('.msp' -contains [System.IO.Path]::GetExtension($_)) })]
        [Alias('Path')]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$AddParameters
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Start-ADTMspProcess]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Start-ADTMspProcess @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Unblock-ADTAppExecution
#
#---------------------------------------------------------------------------

function Unblock-AppExecution
{
    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Unblock-ADTAppExecution]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Unblock-ADTAppExecution
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Block-ADTAppExecution
#
#---------------------------------------------------------------------------

function Block-AppExecution
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, HelpMessage = 'Specify process names, separated by commas.')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ProcessName
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Block-ADTAppExecution]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Block-ADTAppExecution @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTRegistryValue
#
#---------------------------------------------------------------------------

function Test-RegistryValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Key,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Alias('Value')]
        [System.Object]$Name,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce deprecation of function and set up accumulator for all piped in keys.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-ADTRegistryValue]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
        $keys = [System.Collections.Generic.List[System.Object]]::new()
    }

    process
    {
        # Add all keys to the collector.
        $keys.Add($Key)
    }

    end
    {
        # Process provided keys if we have any.
        if ($keys.Count)
        {
            try
            {
                if ($PSBoundParameters.ContainsKey('Key'))
                {
                    $null = $PSBoundParameters.Remove('Key')
                }
                $keys | Test-ADTRegistryValue @PSBoundParameters
            }
            catch
            {
                $PSCmdlet.ThrowTerminatingError($_)
            }
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Convert-ADTRegistryPath
#
#---------------------------------------------------------------------------

function Convert-RegistryPath
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$DisableFunctionLogging = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Convert-ADTRegistryPath]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('DisableFunctionLogging'))
    {
        $null = $PSBoundParameters.Remove('DisableFunctionLogging')
    }
    if (!$DisableFunctionLogging)
    {
        $PSBoundParameters.Add('InformationAction', [System.Management.Automation.ActionPreference]::Continue)
    }
    try
    {
        Convert-ADTRegistryPath @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTMSUpdates
#
#---------------------------------------------------------------------------

function Test-MSUpdates
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Enter the KB Number for the Microsoft Update')]
        [ValidateNotNullOrEmpty()]
        [System.String]$KbNumber,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-ADTMSUpdates]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Test-ADTMSUpdates @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTBattery
#
#---------------------------------------------------------------------------

function Test-Battery
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-ADTBattery]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Test-ADTBattery @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Start-ADTServiceAndDependencies
#
#---------------------------------------------------------------------------

function Start-ServiceAndDependencies
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String]$Service,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ComputerName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipServiceExistsTest,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipDependentServices,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$PendingStatusWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and dead parameters.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Start-ADTServiceAndDependencies]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    $null = ('ComputerName', 'SkipServiceExistsTest').ForEach({
            if ($PSBoundParameters.ContainsKey($_))
            {
                Write-ADTLogEntry -Message "The parameter '-$_' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
                $PSBoundParameters.Remove($_)
            }
        })

    # Translate $ContinueOnError to an ActionPreference before executing.
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    try
    {
        Start-ADTServiceAndDependencies @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Stop-ADTServiceAndDependencies
#
#---------------------------------------------------------------------------

function Stop-ServiceAndDependencies
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String]$Service,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ComputerName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipServiceExistsTest,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipDependentServices,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$PendingStatusWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and dead parameters.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Stop-ADTServiceAndDependencies]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    $null = ('ComputerName', 'SkipServiceExistsTest').ForEach({
            if ($PSBoundParameters.ContainsKey($_))
            {
                Write-ADTLogEntry -Message "The parameter '-$_' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
                $PSBoundParameters.Remove($_)
            }
        })

    # Translate $ContinueOnError to an ActionPreference before executing.
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    try
    {
        Stop-ADTServiceAndDependencies @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTRegistryKey
#
#---------------------------------------------------------------------------

function Set-RegistryKey
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Object]$Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Win32.RegistryValueKind]$Type,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTRegistryKey]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Set-ADTRegistryKey @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Remove-ADTRegistryKey
#
#---------------------------------------------------------------------------

function Remove-RegistryKey
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Remove-ADTRegistryKey]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Remove-ADTRegistryKey @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Remove-ADTFileFromUserProfiles
#
#---------------------------------------------------------------------------

function Remove-FileFromUserProfiles
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Path,

        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$LiteralPath,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ExcludeNTAccount,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExcludeSystemProfiles = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExcludeServiceProfiles = $true,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeDefaultUser,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and dead parameters.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Remove-ADTFileFromUserProfiles]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    $null = ('SystemProfiles', 'ServiceProfiles').Where({ $PSBoundParameters.ContainsKey("Exclude$_") }).ForEach({
            if (!$PSBoundParameters."Exclude$_")
            {
                $PSBoundParameters.Add("Include$_", [System.Management.Automation.SwitchParameter]$true)
            }
            $PSBoundParameters.Remove("Exclude$_")
        })
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        Write-ADTLogEntry -Message "The parameter '-ContinueOnError' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }

    try
    {
        Remove-ADTFileFromUserProfiles @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTRegistryKey
#
#---------------------------------------------------------------------------

function Get-RegistryKey
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Alias('Value')]
        [System.String]$Name = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ReturnEmptyKeyIfExists,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DoNotExpandEnvironmentNames,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTRegistryKey]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Get-ADTRegistryKey @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Install-ADTMSUpdates
#
#---------------------------------------------------------------------------

function Install-MSUpdates
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Directory
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Install-ADTMSUpdates]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Install-ADTMSUpdates @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Direct implementation for Get-SchedulerTask.
#
#---------------------------------------------------------------------------

function Get-SchedulerTask
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'TaskName', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TaskName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Make this function continue on error.
        Initialize-ADTFunction -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction ('Stop', 'SilentlyContinue')[$ContinueOnError]

        # Advise that this function is considered deprecated.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] is deprecated. Please migrate your scripts to use the built-in [Get-ScheduledTask] Cmdlet." -Severity 2
    }

    process
    {
        Write-ADTLogEntry -Message 'Retrieving Scheduled Tasks...'
        try
        {
            try
            {
                # Get CSV data from the binary and confirm success.
                $exeSchtasksResults = & "$([System.Environment]::SystemDirectory)\schtasks.exe" /Query /V /FO CSV 2>&1
                if ($Global:LASTEXITCODE -ne 0)
                {
                    $naerParams = @{
                        Exception = [System.Runtime.InteropServices.ExternalException]::new("The call to [$([System.Environment]::SystemDirectory)\schtasks.exe] failed with exit code [$Global:LASTEXITCODE].", $Global:LASTEXITCODE)
                        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                        ErrorId = 'SchTasksExecutableFailure'
                        TargetObject = $exeSchtasksResults
                        RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                    }
                    throw (New-ADTErrorRecord @naerParams)
                }

                # Convert CSV data to objects and re-process to remove non-word characters before returning data to the caller.
                if (($schTasks = $exeSchtasksResults | ConvertFrom-Csv | & { process { if (($_.TaskName -match '^\\') -and (!$PSBoundParameters.ContainsKey('TaskName') -or ($_.TaskName -match $TaskName))) { return $_ } } }))
                {
                    return $schTasks | Select-Object -Property ($schTasks[0].PSObject.Properties.Name | & {
                            process
                            {
                                @{ Label = $_ -replace '[^\w]'; Expression = [scriptblock]::Create("`$_.'$_'") }
                            }
                        })
                }
            }
            catch
            {
                Write-Error -ErrorRecord $_
            }
        }
        catch
        {
            Invoke-ADTFunctionErrorHandler -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to retrieve scheduled tasks."
        }
    }

    end
    {
        Complete-ADTFunction -Cmdlet $PSCmdlet
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTPendingReboot
#
#---------------------------------------------------------------------------

function Get-PendingReboot
{
    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTPendingReboot]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    try
    {
        Get-ADTPendingReboot
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Invoke-ADTRegSvr32
#
#---------------------------------------------------------------------------

function Invoke-RegisterOrUnregisterDLL
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Register', 'Unregister')]
        [Alias('DLLAction')]
        [System.String]$Action = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Invoke-ADTRegSvr32]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Invoke-ADTRegSvr32 @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Register-ADTDll
#
#---------------------------------------------------------------------------

function Register-DLL
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Register-ADTDll]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Register-ADTDll @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Unregister-ADTDll
#
#---------------------------------------------------------------------------

function Unregister-DLL
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Unregister-ADTDll]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Unregister-ADTDll @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Remove-ADTFolder
#
#---------------------------------------------------------------------------

function Remove-Folder
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DisableRecursion,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Remove-ADTFolder]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Remove-ADTFolder @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTActiveSetup
#
#---------------------------------------------------------------------------

function Set-ActiveSetup
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$StubExePath,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Arguments = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Description = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Version = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Locale = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [System.Management.Automation.SwitchParameter]$DisableActiveSetup,

        [Parameter(Mandatory = $true, ParameterSetName = 'Purge')]
        [System.Management.Automation.SwitchParameter]$PurgeActiveSetupKey,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ExecuteForCurrentUser = $true,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTActiveSetup]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ExecuteForCurrentUser'))
    {
        $PSBoundParameters.Add('NoExecuteForCurrentUser', !$PSBoundParameters.ExecuteForCurrentUser)
        $null = $PSBoundParameters.Remove('ExecuteForCurrentUser')
    }
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if ($StubExePath.EndsWith('.ps1'))
    {
        $PSBoundParameters.Add('ExecutionPolicy', [Microsoft.PowerShell.ExecutionPolicy]::Bypass)
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Set-ADTActiveSetup @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTItemPermission
#
#---------------------------------------------------------------------------

function Set-ItemPermission
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Path to the folder or file you want to modify (ex: C:\Temp)', ParameterSetName = 'DisableInheritance')]
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Path to the folder or file you want to modify (ex: C:\Temp)', ParameterSetName = 'EnableInheritance')]
        [ValidateNotNullOrEmpty()]
        [Alias('File', 'Folder')]
        [System.String]$Path,

        [Parameter( Mandatory = $true, Position = 1, HelpMessage = 'One or more user names (ex: BUILTIN\Users, DOMAIN\Admin). If you want to use SID, prefix it with an asterisk * (ex: *S-1-5-18)', ParameterSetName = 'DisableInheritance')]
        [Alias('Username', 'Users', 'SID', 'Usernames')]
        [System.String[]]$User,

        [Parameter( Mandatory = $true, Position = 2, HelpMessage = "Permission or list of permissions to be set/added/removed/replaced. To see all the possible permissions go to 'http://technet.microsoft.com/fr-fr/library/ff730951.aspx'", ParameterSetName = 'DisableInheritance')]
        [Alias('Acl', 'Grant', 'Permissions', 'Deny')]
        [ValidateSet('AppendData', 'ChangePermissions', 'CreateDirectories', 'CreateFiles', 'Delete', `
                'DeleteSubdirectoriesAndFiles', 'ExecuteFile', 'FullControl', 'ListDirectory', 'Modify', `
                'Read', 'ReadAndExecute', 'ReadAttributes', 'ReadData', 'ReadExtendedAttributes', 'ReadPermissions', `
                'Synchronize', 'TakeOwnership', 'Traverse', 'Write', 'WriteAttributes', 'WriteData', 'WriteExtendedAttributes', 'None')]
        [System.String[]]$Permission,

        [Parameter(Mandatory = $false, Position = 3, HelpMessage = 'Whether you want to set Allow or Deny permissions', ParameterSetName = 'DisableInheritance')]
        [Alias('AccessControlType')]
        [ValidateSet('Allow', 'Deny')]
        [System.String]$PermissionType = 'Allow',

        [Parameter(Mandatory = $false, Position = 4, HelpMessage = 'Sets how permissions are inherited', ParameterSetName = 'DisableInheritance')]
        [ValidateSet('ContainerInherit', 'None', 'ObjectInherit')]
        [System.String[]]$Inheritance = 'None',

        [Parameter(Mandatory = $false, Position = 5, HelpMessage = 'Sets how to propage inheritance flags', ParameterSetName = 'DisableInheritance')]
        [ValidateSet('None', 'InheritOnly', 'NoPropagateInherit')]
        [System.String]$Propagation = 'None',

        [Parameter(Mandatory = $false, Position = 6, HelpMessage = 'Specifies which method will be used to add/remove/replace permissions.', ParameterSetName = 'DisableInheritance')]
        [ValidateSet('Add', 'Set', 'Reset', 'Remove', 'RemoveSpecific', 'RemoveAll')]
        [Alias('ApplyMethod', 'ApplicationMethod')]
        [System.String]$Method = 'Add',

        [Parameter(Mandatory = $true, Position = 1, HelpMessage = 'Enables inheritance, which removes explicit permissions.', ParameterSetName = 'EnableInheritance')]
        [System.Management.Automation.SwitchParameter]$EnableInheritance
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTItemPermission]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('Method'))
    {
        $PSBoundParameters.Method = $PSBoundParameters.Method -replace '^(Add|Set|Reset|Remove)(Specific|All)?$', '$1AccessRule$2'
    }
    try
    {
        Set-ADTItemPermission @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around New-ADTMsiTransform
#
#---------------------------------------------------------------------------

function New-MsiTransform
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$MsiPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ApplyTransformPath = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$NewTransformPath = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]$TransformProperties,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [New-ADTMsiTransform]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        New-ADTMsiTransform @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Invoke-ADTSCCMTask
#
#---------------------------------------------------------------------------

function Invoke-SCCMTask
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('HardwareInventory', 'SoftwareInventory', 'HeartbeatDiscovery', 'SoftwareInventoryFileCollection', 'RequestMachinePolicy', 'EvaluateMachinePolicy', 'LocationServicesCleanup', 'SoftwareMeteringReport', 'SourceUpdate', 'PolicyAgentCleanup', 'RequestMachinePolicy2', 'CertificateMaintenance', 'PeerDistributionPointStatus', 'PeerDistributionPointProvisioning', 'ComplianceIntervalEnforcement', 'SoftwareUpdatesAgentAssignmentEvaluation', 'UploadStateMessage', 'StateMessageManager', 'SoftwareUpdatesScan', 'AMTProvisionCycle', 'UpdateStorePolicy', 'StateSystemBulkSend', 'ApplicationManagerPolicyAction', 'PowerManagementStartSummarizer')]
        [System.String]$ScheduleID,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Invoke-ADTSCCMTask]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Invoke-ADTSCCMTask @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Install-ADTSCCMSoftwareUpdates
#
#---------------------------------------------------------------------------

function Install-SCCMSoftwareUpdates
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$SoftwareUpdatesScanWaitInSeconds,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$WaitForPendingUpdatesTimeout,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Install-ADTSCCMSoftwareUpdates]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Install-ADTSCCMSoftwareUpdates @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Send-ADTKeys
#
#---------------------------------------------------------------------------

function Send-Keys
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowEmptyString()]
        [ValidateNotNull()]
        [System.String]$WindowTitle = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$GetAllWindowTitles,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.IntPtr]$WindowHandle,

        [Parameter(Mandatory = $false, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Keys = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$WaitSeconds
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Send-ADTKeys]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('WaitSeconds'))
    {
        $PSBoundParameters.WaitDuration = [System.TimeSpan]::FromSeconds($WaitSeconds)
        $null = $PSBoundParameters.Remove('WaitSeconds')
    }
    try
    {
        Send-ADTKeys @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTShortcut
#
#---------------------------------------------------------------------------

function Get-Shortcut
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTShortcut]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Get-ADTShortcut @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTShortcut
#
#---------------------------------------------------------------------------

function Set-Shortcut
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, ParameterSetName = 'Pipeline')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]$PathHash,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TargetPath = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Arguments = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IconLocation = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IconIndex = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Description = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Normal', 'Maximized', 'Minimized', 'DontChange')]
        [System.String]$WindowStyle = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Nullable[System.Boolean]]$RunAsAdmin,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Hotkey = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTShortcut]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
        if ($PSBoundParameters.ContainsKey('ContinueOnError'))
        {
            $null = $PSBoundParameters.Remove('ContinueOnError')
        }
        if (!$ContinueOnError)
        {
            $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
        }

        # Set up collector for piped in path objects.
        $paths = [System.Collections.Generic.List[System.String]]::new()
    }

    process
    {
        # Add all paths to the collector.
        if ($PSCmdlet.ParameterSetName.Equals('Default'))
        {
            $paths.Add($Path)
        }
        elseif ($PSCmdlet.ParameterSetName.Equals('Pipeline') -and $PathHash.ContainsKey('Path') -and ![System.String]::IsNullOrWhiteSpace($PathHash.Path))
        {
            $paths.Add($PathHash.Path)
        }
    }

    end
    {
        # Process provided paths if we have any.
        if ($paths.Count)
        {
            try
            {
                if ($PSBoundParameters.ContainsKey('Path'))
                {
                    $null = $PSBoundParameters.Remove('Path')
                }
                $paths | Set-ADTShortcut @PSBoundParameters
            }
            catch
            {
                if (!$ContinueOnError)
                {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            }
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around New-ADTShortcut
#
#---------------------------------------------------------------------------

function New-Shortcut
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TargetPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Arguments = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IconLocation = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$IconIndex,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Description = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Normal', 'Maximized', 'Minimized')]
        [System.String]$WindowStyle = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$RunAsAdmin,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Hotkey = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [New-ADTShortcut]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        New-ADTShortcut @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Start-ADTProcessAsUser
#
#---------------------------------------------------------------------------

function Execute-ProcessAsUser
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = "Silenced to get the module build system going. This function is yet to be refactored.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$UserName = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Path')]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TempPath = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Alias('Parameters')]
        [System.String]$ArgumentList = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [Alias('SecureParameters')]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateSet('HighestAvailable', 'LeastPrivilege')]
        [System.String]$RunLevel = 'HighestAvailable',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$Wait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Start-ADTProcessAsUser]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Announce dead parameters.
    $null = ('TempPath', 'RunLevel').ForEach({
            if ($PSBoundParameters.ContainsKey($_))
            {
                Write-ADTLogEntry -Message "The parameter '-$_' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
                $PSBoundParameters.Remove($_)
            }
        })

    # Translate Wait parameter, which is now NoWait in 4.1
    if ($PSBoundParameters.ContainsKey('Wait'))
    {
        $PSBoundParameters.Add('NoWait', !$PSBoundParameters.Wait)
        $null = $PSBoundParameters.Remove('Wait')
    }
    else
    {
        $PSBoundParameters.Add('NoWait', $true)
    }

    # Translate the ContinueOnError state.
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    # Invoke underlying function.
    try
    {
        if (($res = Start-ADTProcessAsUser @PSBoundParameters) -and $PassThru)
        {
            return $res.Result
        }
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Close-ADTInstallationProgress
#
#---------------------------------------------------------------------------

function Close-InstallationProgress
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 60)]
        [System.Int32]$WaitingTime = 5
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Close-ADTInstallationProgress]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('WaitingTime'))
    {
        Write-ADTLogEntry -Message "The parameter '-WaitingTime' is discontinued and no longer has any effect." -Severity 2 -Source $MyInvocation.MyCommand.Name
    }

    # Invoke underlying function.
    try
    {
        Close-ADTInstallationProgress
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around ConvertTo-ADTNTAccountOrSID
#
#---------------------------------------------------------------------------

function ConvertTo-NTAccountOrSID
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'NTAccountToSID', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$AccountName,

        [Parameter(Mandatory = $true, ParameterSetName = 'SIDToNTAccount', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID,

        [Parameter(Mandatory = $true, ParameterSetName = 'WellKnownName', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WellKnownSIDName,

        [Parameter(Mandatory = $false, ParameterSetName = 'WellKnownName')]
        [System.Management.Automation.SwitchParameter]$WellKnownToNTAccount
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce overall deprecation and any dead parameters before executing.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [ConvertTo-ADTNTAccountOrSID]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

        # Set up collector for pipelined input.
        $pipedInput = [System.Collections.Generic.List[System.String]]::new()
    }

    process
    {
        # Only add non-null strings to our collector.
        if (![System.String]::IsNullOrWhiteSpace(($thisInput = Get-Variable -Name $PSCmdlet.ParameterSetName -ValueOnly)))
        {
            $pipedInput.Add($thisInput)
        }
    }

    end
    {
        # Only proceed if we have collected input.
        if (!$pipedInput.Count)
        {
            return
        }

        try
        {
            $null = $PSBoundParameters.Remove($PSCmdlet.ParameterSetName)
            $pipedInput | ConvertTo-ADTNTAccountOrSID @PSBoundParameters
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTDeferHistory
#
#---------------------------------------------------------------------------

function Get-DeferHistory
{
    [CmdletBinding()]
    param
    (
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTDeferHistory]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Invoke underlying function.
    try
    {
        Get-ADTDeferHistory
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTDeferHistory
#
#---------------------------------------------------------------------------

function Set-DeferHistory
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$DeferTimesRemaining,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [System.String]$DeferDeadline
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTDeferHistory]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Invoke underlying function.
    try
    {
        Set-ADTDeferHistory @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTMsiTableProperty
#
#---------------------------------------------------------------------------

function Get-MsiTableProperty
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$TransformPath,

        [Parameter(Mandatory = $false, ParameterSetName = 'TableInfo')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Table = [System.Management.Automation.Language.NullString]::Value,

        [Parameter(Mandatory = $false, ParameterSetName = 'TableInfo')]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$TablePropertyNameColumnNum,

        [Parameter(Mandatory = $false, ParameterSetName = 'TableInfo')]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$TablePropertyValueColumnNum,

        [Parameter(Mandatory = $true, ParameterSetName = 'SummaryInfo')]
        [System.Management.Automation.SwitchParameter]$GetSummaryInformation,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTMsiTableProperty]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Get-ADTMsiTableProperty @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Set-ADTMsiProperty
#
#---------------------------------------------------------------------------

function Set-MsiProperty
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.__ComObject]$DataBase,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$PropertyName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$PropertyValue,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Set-ADTMsiProperty]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }
    try
    {
        Set-ADTMsiProperty @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTMsiExitCodeMessage
#
#---------------------------------------------------------------------------

function Get-MsiExitCodeMessage
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Nullable[System.Int32]]$MsiExitCode
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTMsiExitCodeMessage]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Invoke underlying function.
    try
    {
        Get-ADTMsiExitCodeMessage @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTObjectProperty
#
#---------------------------------------------------------------------------

function Get-ObjectProperty
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$InputObject,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]$PropertyName,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.Object[]]$ArgumentList
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTObjectProperty]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Invoke underlying function.
    try
    {
        Get-ADTObjectProperty @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Invoke-ADTObjectMethod
#
#---------------------------------------------------------------------------

function Invoke-ObjectMethod
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$InputObject,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]$MethodName,

        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = 'Positional')]
        [ValidateNotNullOrEmpty()]
        [System.Object[]]$ArgumentList,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'Named')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]$Parameter
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Invoke-ADTObjectMethod]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Invoke underlying function.
    try
    {
        Invoke-ADTObjectMethod @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Get-ADTPEFileArchitecture
#
#---------------------------------------------------------------------------

function Get-PEFileArchitecture
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
        [Systemn.IO.FileInfo[]]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Systemn.Boolean]$ContinueOnError = $true,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        # Set strict mode to the highest within this function's scope.
        Set-StrictMode -Version 3

        # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
        Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTPEFileArchitecture]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
        if ($PSBoundParameters.ContainsKey('ContinueOnError'))
        {
            $null = $PSBoundParameters.Remove('ContinueOnError')
        }
        if (!$ContinueOnError)
        {
            $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
        }

        # Set up collector for pipelined input.
        $filePaths = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    }

    process
    {
        # Collect all input for processing at the end.
        if ($null -ne $FilePath)
        {
            $filePaths.Add($FilePath)
        }
    }

    end
    {
        # Only process if we have files in our collector.
        if (!$filePaths.Count)
        {
            return
        }

        try
        {
            if ($PSBoundParameters.ContainsKey('FilePath'))
            {
                $null = $PSBoundParameters.Remove('FilePath')
            }
            $filePaths | Get-ADTPEFileArchitecture @PSBoundParameters | & {
                process
                {
                    switch ([System.UInt16]$_)
                    {
                        0
                        {
                            # The contents of this file are assumed to be applicable to any machine type
                            'Native'
                            break
                        }
                        0x014C
                        {
                            # File for Windows 32-bit systems
                            '32BIT'
                            break
                        }
                        0x0200
                        {
                            # File for Intel Itanium x64 processor family
                            'Itanium-x64'
                            break
                        }
                        0x8664
                        {
                            # File for Windows 64-bit systems
                            '64BIT'
                            break
                        }
                        default
                        {
                            'Unknown'
                            break
                        }
                    }

                }
            }
        }
        catch
        {
            if (!$ContinueOnError)
            {
                $PSCmdlet.ThrowTerminatingError($_)
            }
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around Test-ADTMutexAvailability
#
#---------------------------------------------------------------------------

function Test-IsMutexAvailable
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateLength(1, 260)]
        [System.String]$MutexName,

        [Parameter(Mandatory = $false)]
        [ValidateScript({ ($_ -ge -1) -and ($_ -le [System.Int32]::MaxValue) })]
        [System.Nullable[System.Int32]]$MutexWaitTimeInMilliseconds = 1
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and any dead parameters before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Test-ADTMutexAvailability]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings

    # Convert out changed parameter.
    if ($PSBoundParameters.ContainsKey('MutexWaitTimeInMilliseconds'))
    {
        $PSBoundParameters.MutexWaitTime = [System.TimeSpan]::FromMilliseconds($MutexWaitTimeInMilliseconds)
        $null = $PSBoundParameters.Remove('MutexWaitTimeInMilliseconds')
    }

    # Invoke underlying function.
    try
    {
        Test-ADTMutexAvailability @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#---------------------------------------------------------------------------
#
# MARK: Wrapper around New-ADTZipFile
#
#---------------------------------------------------------------------------

function New-ZipFile
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DestinationArchiveDirectoryPath,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DestinationArchiveFileName,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SourceDirectoryPath')]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
        [System.String[]]$SourceDirectoryPath,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SourceFilePath')]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
        [System.String[]]$SourceFilePath,

        [Parameter(Mandatory = $false, Position = 3)]
        [System.Management.Automation.SwitchParameter]$RemoveSourceAfterArchiving,

        [Parameter(Mandatory = $false, Position = 4)]
        [System.Management.Automation.SwitchParameter]$OverWriteArchive,

        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.Boolean]$ContinueOnError = $true
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce overall deprecation and translate $ContinueOnError to an ActionPreference before executing.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [New-ADTZipFile]. Please migrate your scripts to use the new function." -Severity 2 -DebugMessage:$noDepWarnings
    if ($PSBoundParameters.ContainsKey('ContinueOnError'))
    {
        $null = $PSBoundParameters.Remove('ContinueOnError')
    }
    if (!$ContinueOnError)
    {
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::Stop
    }

    # Convert source path parameter.
    $PSBoundParameters.Add('LiteralPath', $PSBoundParameters.($PSCmdlet.ParameterSetName))
    $null = $PSBoundParameters.Remove($PSCmdlet.ParameterSetName)

    # Convert destination parameters.
    $PSBoundParameters.Add('DestinationPath', (Join-Path -Path $DestinationArchiveDirectoryPath -ChildPath $DestinationArchiveFileName))
    $null = $PSBoundParameters.Remove('DestinationArchiveDirectoryPath')
    $null = $PSBoundParameters.Remove('DestinationArchiveFileName')

    # Convert $OverWriteArchive.
    if ($PSBoundParameters.ContainsKey('OverWriteArchive'))
    {
        $PSBoundParameters.Add('Force', $OverWriteArchive)
        $null = $PSBoundParameters.Remove('OverWriteArchive')
    }

    # Invoke replacement function.
    try
    {
        New-ADTZipFile @PSBoundParameters
    }
    catch
    {
        if (!$ContinueOnError)
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}


#---------------------------------------------------------------------------
#
# MARK: Deprecation announcement for Set-PinnedApplication
#
#---------------------------------------------------------------------------

function Set-PinnedApplication
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This compatibility wrapper function cannot support ShouldProcess for backwards compatiblity purposes.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Action', Justification = "The parameter is not used as the function is a deprecation announcement and performs no actions.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'FilePath', Justification = "The parameter is not used as the function is a deprecation announcement and performs no actions.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('PinToStartMenu', 'UnpinFromStartMenu', 'PinToTaskbar', 'UnpinFromTaskbar')]
        [System.String]$Action,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    # Announce that this function is no more and therefore does nothing within the deployment script.
    Write-ADTLogEntry -Message "The function [$($MyInvocation.MyCommand.Name)] has been removed from PSAppDeployToolkit as its functionality no longer works with Windows 10 1809 or higher targets." -Severity 2
}


#---------------------------------------------------------------------------
#
# MARK: Direct copy of Write-FunctionHeaderOrFooter for backwards compatibility reasons.
#
#---------------------------------------------------------------------------

function Write-FunctionHeaderOrFooter
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$CmdletName,

        [Parameter(Mandatory = $true, ParameterSetName = 'Header')]
        [AllowEmptyCollection()]
        [System.Collections.Generic.IReadOnlyDictionary[System.String, System.Object]]$CmdletBoundParameters,

        [Parameter(Mandatory = $true, ParameterSetName = 'Header')]
        [System.Management.Automation.SwitchParameter]$Header,

        [Parameter(Mandatory = $true, ParameterSetName = 'Footer')]
        [System.Management.Automation.SwitchParameter]$Footer
    )

    # Set strict mode to the highest within this function's scope.
    Set-StrictMode -Version 3

    if ($Header)
    {
        Write-ADTLogEntry -Message 'Function Start' -Source ${CmdletName} -DebugMessage

        # Get the parameters that the calling function was invoked with.
        if ([System.String]$CmdletBoundParameters = $CmdletBoundParameters | Format-Table -Property @{ Label = 'Parameter'; Expression = { "[-$($_.Key)]" } }, @{ Label = 'Value'; Expression = { $_.Value }; Alignment = 'Left' }, @{ Label = 'Type'; Expression = { $_.Value.GetType().Name }; Alignment = 'Left' } -AutoSize -Wrap | Out-String)
        {
            Write-ADTLogEntry -Message "Function invoked with bound parameter(s): `r`n$CmdletBoundParameters" -Source ${CmdletName} -DebugMessage
        }
        else
        {
            Write-ADTLogEntry -Message 'Function invoked without any bound parameters.' -Source ${CmdletName} -DebugMessage
        }
    }
    elseif ($Footer)
    {
        Write-ADTLogEntry -Message 'Function End' -Source ${CmdletName} -DebugMessage
    }
}


#---------------------------------------------------------------------------
#
# MARK: Module and session code
#
#---------------------------------------------------------------------------

# Set required variables to ensure module functionality.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 3

# Import our module backend.
$adtModule = if (Test-Path -LiteralPath "$PSScriptRoot\PSAppDeployToolkit" -PathType Container)
{
    Get-ChildItem -LiteralPath $PSScriptRoot\PSAppDeployToolkit -Recurse -File | Unblock-File
    Import-Module -FullyQualifiedName @{ ModuleName = "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force -PassThru -ErrorAction Stop
}
elseif (Test-Path -LiteralPath "$PSScriptRoot\..\..\..\..\PSAppDeployToolkit" -PathType Container)
{
    Get-ChildItem -LiteralPath $PSScriptRoot\..\..\..\..\PSAppDeployToolkit -Recurse -File | Unblock-File
    Import-Module -FullyQualifiedName @{ ModuleName = "$PSScriptRoot\..\..\..\..\PSAppDeployToolkit\PSAppDeployToolkit.psd1"; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force -PassThru -ErrorAction Stop
}
else
{
    Import-Module -FullyQualifiedName @{ ModuleName = 'PSAppDeployToolkit'; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force -PassThru -ErrorAction Stop
}

# Build out parameter hashtable.
$sessionParams = @{}; $adtModule.ExportedCommands.'Open-ADTSession'.Parameters.Values | & {
    process
    {
        # Skip any Open-ADTSession params that are considered frontend params/variables.
        if ($_.ParameterSets.Values.HelpMessage -notmatch '^Frontend (Parameter|Variable)$')
        {
            return
        }

        # Skip if we didn't get a variable value.
        if (!($variable = Get-Variable -Name $_.Name -ErrorAction Ignore))
        {
            return
        }

        # Add the parameter if it's not null.
        if (![System.String]::IsNullOrWhiteSpace((Out-String -InputObject $variable.Value)))
        {
            $sessionParams.Add($variable.Name, $variable.Value)
        }
    }
}

# Remove dates if they fail to parse using local culture settings.
if ($sessionParams.ContainsKey('AppScriptDate'))
{
    try
    {
        $sessionParams.AppScriptDate = [System.DateTime]::Parse($sessionParams.AppScriptDate, $Host.CurrentCulture)
    }
    catch
    {
        $null = $sessionParams.Remove('AppScriptDate')
    }
}

# Redefine DeployAppScriptParameters due bad casting in Deploy-Application.ps1.
if ($sessionParams.ContainsKey('DeployAppScriptParameters'))
{
    $sessionParams.DeployAppScriptParameters = (Get-PSCallStack)[1].InvocationInfo.BoundParameters
}

# Open a new deployment session.
Open-ADTSession -SessionState $ExecutionContext.SessionState @sessionParams

# Define aliases for some functions to maintain backwards compatibility.
New-Alias -Name Refresh-SessionEnvironmentVariables -Value Update-SessionEnvironmentVariables -Option ReadOnly -Force
New-Alias -Name Refresh-Desktop -Value Update-Desktop -Option ReadOnly -Force

# Finalize setup of AppDeployToolkitMain.ps1.
Set-Item -LiteralPath $adtWrapperFuncs -Options ReadOnly
New-Variable -Name noDepWarnings -Value (($adtConfig = Get-ADTConfig).Toolkit.ContainsKey('WrapperWarnings') -and !$adtConfig.Toolkit.WrapperWarnings) -Option ReadOnly -Force
Remove-Variable -Name adtConfig, adtWrapperFuncs, sessionParams, adtModule -Force -Confirm:$false
Set-StrictMode -Version 1


#---------------------------------------------------------------------------
#
# MARK: Compatibility extension support
#
#---------------------------------------------------------------------------

if ((Test-Path -LiteralPath "$PSScriptRoot\AppDeployToolkitExtensions.ps1" -PathType Leaf))
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'scriptParentPath', Justification = "This variable is used within a dot-sourced script that PSScriptAnalyzer has no visibility of.")]
    $scriptParentPath = if ($invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName)
    {
        # If this script was invoked by another script.
        (Get-Item -LiteralPath $invokingScript).DirectoryName
    }
    else
    {
        # If this script was not invoked by another script, fall back to the directory one level above this script.
        (Get-Item -LiteralPath $MyInvocation.MyCommand.Definition).Directory.Parent.FullName
    }
    . "$PSScriptRoot\AppDeployToolkitExtensions.ps1"
}

# SIG # Begin signature block
# MIIuaQYJKoZIhvcNAQcCoIIuWjCCLlYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCwcuworb6PPsdb
# COjno7m6180gXyeYU8BJJX7krR/6H6CCE5UwggWQMIIDeKADAgECAhAFmxtXno4h
# MuI5B72nd3VcMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0xMzA4MDExMjAwMDBaFw0z
# ODAxMTUxMjAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjQjBAMA8GA1Ud
# EwEB/wQFMAMBAf8wDgYDVR0PAQH/BAQDAgGGMB0GA1UdDgQWBBTs1+OC0nFdZEzf
# Lmc/57qYrhwPTzANBgkqhkiG9w0BAQwFAAOCAgEAu2HZfalsvhfEkRvDoaIAjeNk
# aA9Wz3eucPn9mkqZucl4XAwMX+TmFClWCzZJXURj4K2clhhmGyMNPXnpbWvWVPjS
# PMFDQK4dUPVS/JA7u5iZaWvHwaeoaKQn3J35J64whbn2Z006Po9ZOSJTROvIXQPK
# 7VB6fWIhCoDIc2bRoAVgX+iltKevqPdtNZx8WorWojiZ83iL9E3SIAveBO6Mm0eB
# cg3AFDLvMFkuruBx8lbkapdvklBtlo1oepqyNhR6BvIkuQkRUNcIsbiJeoQjYUIp
# 5aPNoiBB19GcZNnqJqGLFNdMGbJQQXE9P01wI4YMStyB0swylIQNCAmXHE/A7msg
# dDDS4Dk0EIUhFQEI6FUy3nFJ2SgXUE3mvk3RdazQyvtBuEOlqtPDBURPLDab4vri
# RbgjU2wGb2dVf0a1TD9uKFp5JtKkqGKX0h7i7UqLvBv9R0oN32dmfrJbQdA75PQ7
# 9ARj6e/CVABRoIoqyc54zNXqhwQYs86vSYiv85KZtrPmYQ/ShQDnUBrkG5WdGaG5
# nLGbsQAe79APT0JsyQq87kP6OnGlyE0mpTX9iV28hWIdMtKgK1TtmlfB2/oQzxm3
# i0objwG2J5VT6LaJbVu8aNQj6ItRolb58KaAoNYes7wPD1N1KarqE3fk3oyBIa0H
# EEcRrYc9B9F1vM/zZn4wggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0G
# CSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTla
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C
# 0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce
# 2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0da
# E6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6T
# SXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoA
# FdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7Oh
# D26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM
# 1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z
# 8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05
# huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNY
# mtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP
# /2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0T
# AQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYD
# VR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNV
# HR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATAN
# BgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95Ry
# sQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HL
# IvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5Btf
# Q/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnh
# OE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIh
# dXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV
# 9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/j
# wVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYH
# Ki8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmC
# XBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l
# /aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZW
# eE4wggdJMIIFMaADAgECAhAK+Vu2vqIMhQ6YxvuOrAj5MA0GCSqGSIb3DQEBCwUA
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwHhcNMjQwOTA1MDAwMDAwWhcNMjcwOTA3MjM1OTU5WjCB0TET
# MBEGCysGAQQBgjc8AgEDEwJVUzEZMBcGCysGAQQBgjc8AgECEwhDb2xvcmFkbzEd
# MBsGA1UEDwwUUHJpdmF0ZSBPcmdhbml6YXRpb24xFDASBgNVBAUTCzIwMTMxNjM4
# MzI3MQswCQYDVQQGEwJVUzERMA8GA1UECBMIQ29sb3JhZG8xFDASBgNVBAcTC0Nh
# c3RsZSBSb2NrMRkwFwYDVQQKExBQYXRjaCBNeSBQQywgTExDMRkwFwYDVQQDExBQ
# YXRjaCBNeSBQQywgTExDMIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEA
# uydxko2Hrl6sANJUjfdypKP60qBH5EkhfaRQAnn+e3vg2eVcbiEWIjlrMYzvK2sg
# OMBbwGebqAURkFmUCKDdGxcxKeuXdaXPHWPKwc2WjYCFajrX6HofiiwNzOCdL6VE
# 4PDQhPRR7SIdNNFSrx5C4ZDN1T6OH+ydX7EQF8+NBUNHRbEVdl+h9H5Aexx63afa
# 8zu3g/GXluyXKbb+JHtgNJaUgFuFORTxw1TO6qH+S6Hrppf9QcAFmu4xGtkc2FSh
# gv0NgWMNGDZqJr/o9sqJ2tdaZHDyr6H8PvY8egoUshF7ccgEYtEEdB9SRR8mVQik
# 1w5oGTjDWjHj+8jgTpzletRywptk/m8PehVBN8ntqoSdvLLcuQVzmuPLzN/iuKh5
# sZeWvqPONApcEnZcONpXebyiUPnEePr5rZAU7hMjMw2ZPnQlMcbGvtgP2qi7m2f3
# mXFYxWjlKCxaApYHeqSFeWC8zM7OYL2HlZ+GuK4XG8jKVE6sWSW9Wk/dm0vJbasv
# AgMBAAGjggICMIIB/jAfBgNVHSMEGDAWgBRoN+Drtjv4XxGG+/5hewiIZfROQjAd
# BgNVHQ4EFgQU5GCU3SEqeIbhhY9eyU0LcTI75X8wPQYDVR0gBDYwNDAyBgVngQwB
# AzApMCcGCCsGAQUFBwIBFhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwDgYD
# VR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMIG1BgNVHR8Ega0wgaow
# U6BRoE+GTWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3JsMFOgUaBPhk1odHRw
# Oi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmlu
# Z1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNybDCBlAYIKwYBBQUHAQEEgYcwgYQwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBcBggrBgEFBQcwAoZQ
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29k
# ZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcnQwCQYDVR0TBAIwADANBgkq
# hkiG9w0BAQsFAAOCAgEAqA0ub/ilMgdIvMiBeWBoiMxe5OIblObGI7lemcP2WEqa
# EASW11/wVwJU63ZwhtkQaNU4rXjf6fqy5pOUzpQXgYjSaO4D/AOMJKHlypxslFqZ
# /dYpcue2xE3H7lmO4KPf8VxXuFIUqjLetU+kkh7o/Q52RabVAuOrPFKnObixy1HI
# x0/5F+RuP9xhqmDbfM7l5zUAcuOCCkY7buuInEsip9BZXUiVb8K5bPR9Rk7Doat4
# FQmN72xjakcEZOMU/vg0ZgVa8nxkBXtVsjxbsr+bODn0cddHK1QHWil/PmpANkxN
# 7H8tdCAZ8bTzIvvudxSLnt7ssbbQDkAyNw0btDH+MKv/l+VcYyQH51Z5xT9DvHCm
# Ed774boZkP2GfTFvn7/gISEjTdOuUGstdrgSwg1zJPqgK7zWxK48xC7awpa3gwOs
# 9pnyiqHG3rx84/SHUiAL2lkljsD3epmRxsWeZhZNY93xEpQHe9LBvo/t4VRjZzqU
# z+pfEMPqeX/g5+mpb4ap6ZmNJuAYJFmU0LIkCLQN9mKXi1Il9WU6ifn3vYutGMSL
# /BdeWP+7fM7MZLiO+1BIsBdSmV6pZVS3LRBAy3wIlbWL69mvyLCPIQ7z4dtfuzwC
# 36E9k2vhzeiDQ+k1dFJDSdxTDetsck0FuD1ovhiu2caL4BdFsCWsXPLMyvu6OlYx
# ghoqMIIaJgIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEAr5W7a+ogyFDpjG+46sCPkwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgZnrugFjxLVu1TLatedE9qQgrZe8JzDLGqhrro9gGklAw
# DQYJKoZIhvcNAQEBBQAEggGAfk2BR+Yf2nR1+DzU4WVmfRtN/RaUEHniB1tvB7eg
# dJhRxlWoQuVYBmnhNhoyTRusNFoqJHPbyRnASfryk20GpMOtYneUnSWH49Z3TRgq
# eO1B/jqIZ7uVNKkpElGGmJPBXoRHFe5UWrl/3lsZS8B0Wnj3FXy9lWsEVBz+Ilqy
# YMdu2O1GO7Yky0W74BTXOlUYJJAJ1PHdtSB415xLG7FcPBqlQGhaUUS9O1OcmTIG
# Ilinb2h1erdxIpY6jOt7G/B6pouScJ9XHsDcwAR2/iabp2ltACsI73jWcCZ7G5Uj
# zGbR6lpUzuraEXOqG2/SrAFJor4jLwbgX2EQWB59mmYxOckZYLgYNoRZpmMjUUKD
# +BiKEw/uXmJZhpTXiLr3sdkfK+wYKv6cKIg3Ww5YowORyhpMDR0CsUgNY0TJYJWu
# 8Ep+psxgZ55xaKTcxV1BvVODDRPLRkfHknQaGMX4DywczUromPWVtaSQCz9LyaK/
# C4mRedyGfYk73IbtQHHEp2CfoYIXdzCCF3MGCisGAQQBgjcDAwExghdjMIIXXwYJ
# KoZIhvcNAQcCoIIXUDCCF0wCAQMxDzANBglghkgBZQMEAgEFADB4BgsqhkiG9w0B
# CRABBKBpBGcwZQIBAQYJYIZIAYb9bAcBMDEwDQYJYIZIAWUDBAIBBQAEIH89kFr9
# TeoFaukDKRdcll9QmVg8js+ImFs0A6mmwZaFAhEAsxJkwrvqtlFPIpUPSBJHjRgP
# MjAyNjAxMTQxMDU2MzlaoIITOjCCBu0wggTVoAMCAQICEAqA7xhLjfEFgtHEdqeV
# dGgwDQYJKoZIhvcNAQELBQAwaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTAeFw0yNTA2MDQwMDAwMDBaFw0z
# NjA5MDMyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgU0hBMjU2IFJTQTQwOTYgVGltZXN0YW1w
# IFJlc3BvbmRlciAyMDI1IDEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDQRqwtEsae0OquYFazK1e6b1H/hnAKAd/KN8wZQjBjMqiZ3xTWcfsLwOvRxUwX
# cGx8AUjni6bz52fGTfr6PHRNv6T7zsf1Y/E3IU8kgNkeECqVQ+3bzWYesFtkepEr
# vUSbf+EIYLkrLKd6qJnuzK8Vcn0DvbDMemQFoxQ2Dsw4vEjoT1FpS54dNApZfKY6
# 1HAldytxNM89PZXUP/5wWWURK+IfxiOg8W9lKMqzdIo7VA1R0V3Zp3DjjANwqAf4
# lEkTlCDQ0/fKJLKLkzGBTpx6EYevvOi7XOc4zyh1uSqgr6UnbksIcFJqLbkIXIPb
# cNmA98Oskkkrvt6lPAw/p4oDSRZreiwB7x9ykrjS6GS3NR39iTTFS+ENTqW8m6TH
# uOmHHjQNC3zbJ6nJ6SXiLSvw4Smz8U07hqF+8CTXaETkVWz0dVVZw7knh1WZXOLH
# gDvundrAtuvz0D3T+dYaNcwafsVCGZKUhQPL1naFKBy1p6llN3QgshRta6Eq4B40
# h5avMcpi54wm0i2ePZD5pPIssoszQyF4//3DoK2O65Uck5Wggn8O2klETsJ7u8xE
# ehGifgJYi+6I03UuT1j7FnrqVrOzaQoVJOeeStPeldYRNMmSF3voIgMFtNGh86w3
# ISHNm0IaadCKCkUe2LnwJKa8TIlwCUNVwppwn4D3/Pt5pwIDAQABo4IBlTCCAZEw
# DAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQU5Dv88jHt/f3X85FxYxlQQ89hjOgwHwYD
# VR0jBBgwFoAU729TSunkBnx6yuKQVvYv1Ensy04wDgYDVR0PAQH/BAQDAgeAMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMIGVBggrBgEFBQcBAQSBiDCBhTAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMF0GCCsGAQUFBzAChlFodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3Rh
# bXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcnQwXwYDVR0fBFgwVjBUoFKgUIZO
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAZSqt8RwnBLmuYEHs
# 0QhEnmNAciH45PYiT9s1i6UKtW+FERp8FgXRGQ/YAavXzWjZhY+hIfP2JkQ38U+w
# tJPBVBajYfrbIYG+Dui4I4PCvHpQuPqFgqp1PzC/ZRX4pvP/ciZmUnthfAEP1HSh
# TrY+2DE5qjzvZs7JIIgt0GCFD9ktx0LxxtRQ7vllKluHWiKk6FxRPyUPxAAYH2Vy
# 1lNM4kzekd8oEARzFAWgeW3az2xejEWLNN4eKGxDJ8WDl/FQUSntbjZ80FU3i54t
# px5F/0Kr15zW/mJAxZMVBrTE2oi0fcI8VMbtoRAmaaslNXdCG1+lqvP4FbrQ6IwS
# BXkZagHLhFU9HCrG/syTRLLhAezu/3Lr00GrJzPQFnCEH1Y58678IgmfORBPC1JK
# kYaEt2OdDh4GmO0/5cHelAK2/gTlQJINqDr6JfwyYHXSd+V08X1JUPvB4ILfJdmL
# +66Gp3CSBXG6IwXMZUXBhtCyIaehr0XkBoDIGMUG1dUtwq1qmcwbdUfcSYCn+Own
# cVUXf53VJUNOaMWMts0VlRYxe5nK+At+DI96HAlXHAL5SlfYxJ7La54i71McVWRP
# 66bW+yERNpbJCjyCYG2j+bdpxo/1Cy4uPcU3AWVPGrbn5PhDBf3Froguzzhk++am
# i+r3Qrx5bIbY3TVzgiFI7Gq3zWcwgga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIM
# OkmGMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERp
# Z2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQy
# MzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFB
# MD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5
# NiBTSEEyNTYgMjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQC0eDHTCphBcr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq
# +RuwOnPhof6pvF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyF
# Q/4Bt0mAxAHeHYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOE
# CS1UkxBvMgEdgkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdg
# lTcaarps0wjUjsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZ
# a/zbCclF83bRVFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLr
# fxnGpTXiUOeSLsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy
# 3tW/AMOMCZIVNSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJW
# nY0v5ydPpOjL6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdg
# JMoiwOrUG2ZdSoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJ
# SjNm2qA+sdFUeEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkw
# EgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ens
# y04wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNy
# dDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglg
# hkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSm
# f83Qh8WIGjB/T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83a
# fsl3YTj+IQhQE7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5
# nWMQwr8Myb9rEVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ
# 2iA/wdG2th9y1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu
# 4y81hjajV/gxdEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgc
# GV00TYr2Lr3ty9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+
# qNNjqFzeGxcytL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll
# 38F0cuJG7uEBYTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66R
# zIg9sC+NJpud/v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDp
# MPx0LckTetiSuEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8Wh
# baM2tszWkPZPubdcMIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkq
# hkiG9w0BAQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5
# WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJv
# b3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1K
# PDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2r
# snnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C
# 8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBf
# sXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGY
# QJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8
# rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaY
# dj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+
# wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw
# ++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+N
# P8m800ERElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7F
# wI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUw
# AwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAU
# Reuir/SSy4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEB
# BG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsG
# AQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAow
# CDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/
# Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLe
# JLxSA8hO0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE
# 1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9Hda
# XFSMb++hUD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbO
# byMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMYID
# fDCCA3gCAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIElu
# Yy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJT
# QTQwOTYgU0hBMjU2IDIwMjUgQ0ExAhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFl
# AwQCAQUAoIHRMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0B
# CQUxDxcNMjYwMTE0MTA1NjM5WjArBgsqhkiG9w0BCRACDDEcMBowGDAWBBTdYjCs
# hgotMGvaOLFoeVIwB/tBfjAvBgkqhkiG9w0BCQQxIgQgRK35TLY30uxFc2F9LC+X
# vGWIFMQX5vVcGxaSs4h/0XUwNwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQgSqA/oizX
# XITFXJOPgo5na5yuyrM/420mmqM08UYRCjMwDQYJKoZIhvcNAQEBBQAEggIAzrhg
# oQVwAXebvCtlxe+8pILgviwiV0SHtX4toHVVkU9nlLejc19AQF/5xWFQDpoxQ7XS
# IY3QhRG9lmOlk1YC+wZfTA5IQLqZYsMcdaupV+Tioo138xM2pwlK8+YZj0ydDw9H
# iMyEwrWIQzab5HDnKSAzJ2FaCtZp1fWqevtetdB0I2HzY4hyO2Ksxn2kZ9pGjrNv
# wAdH7mmTFvcFEP+0BIv+uaYWB20ApLExhBtl4h5wDE+/DhNJIZkRH6pvTk5UCYPB
# lBd6kZ18qvE5sLBuaW0rBqQMdCgc6DI9Fhw+SG7BSHvqxAz/hwc02e5X5aExLJMD
# CqrQfC5ao2nDoMvLRq9GZrOtSv6r1qYn3vpQVJo7vsxXSGmxmjjgA8XWcp+HRhx3
# Y+9Wk2VzHgn7JdG5AN9civbE1HF8KzPCnlNUbCT3SLZX/ohzipTdh/SBeRKLk70E
# kWTuQEXmm271+MbkqkB7TF5P+QIPYqoVOKU5g8rFYMq/YmkbyLMUnJtxRo2gVQg8
# EleE3f/8zpr7TBpQrd86lFmpYjIPUYIWb8BukNkggPLDqVNqqebShly+bjfmMoqJ
# iXpT1QqH6XwzbLDcXb1v5Ek8RnuUha9AHZEyMwtX0Vm7nG3pww+be+gp78e0V3kI
# Tk5u+Ab1UmE35LzO11kGPo9V/ZCqelYcDvlkLGs=
# SIG # End signature block
