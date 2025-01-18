<#

.SYNOPSIS
PSAppDeployToolkit - This module script contains the PSADT core runtime and functions using by a Invoke-AppDeployToolkit.ps1 script.

.DESCRIPTION
This module can be directly imported from the command line via Import-Module, but it is usually imported by the Invoke-AppDeployToolkit.ps1 script.

This module can usually be updated to the latest version without impacting your per-application Invoke-AppDeployToolkit.ps1 scripts. Please check release notes before upgrading.

PSAppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.LINK
https://psappdeploytoolkit.com

#>

#-----------------------------------------------------------------------------
#
# MARK: Module Initialization Code
#
#-----------------------------------------------------------------------------

# Throw if this psm1 file isn't being imported via our manifest.
if (!([System.Environment]::StackTrace.Split("`n").Trim() -like '*Microsoft.PowerShell.Commands.ModuleCmdletBase.LoadModuleManifest(*'))
{
    throw [System.Management.Automation.ErrorRecord]::new(
        [System.InvalidOperationException]::new("This module must be imported via its .psd1 file, which is recommended for all modules that supply a .psd1 file."),
        'ModuleImportError',
        [System.Management.Automation.ErrorCategory]::InvalidOperation,
        $MyInvocation.MyCommand.ScriptBlock.Module
    )
}

# Clock when the module import starts so we can track it.
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ModuleImportStart', Justification = "This variable is used within ImportsLast.ps1 and therefore cannot be seen here.")]
$ModuleImportStart = [System.DateTime]::Now

# Build out lookup table for all cmdlets used within module, starting with the core cmdlets.
$CommandTable = [System.Collections.Generic.Dictionary[System.String, System.Management.Automation.CommandInfo]]::new()
$ExecutionContext.SessionState.InvokeCommand.GetCmdlets() | & { process { if ($_.PSSnapIn -and $_.PSSnapIn.Name.Equals('Microsoft.PowerShell.Core') -and $_.PSSnapIn.IsDefault) { $CommandTable.Add($_.Name, $_) } } }

# Expand command lookup table with cmdlets used through this module.
& {
    $RequiredModules = [System.Collections.ObjectModel.ReadOnlyCollection[Microsoft.PowerShell.Commands.ModuleSpecification]]$(
        @{ ModuleName = 'CimCmdlets'; Guid = 'fb6cc51d-c096-4b38-b78d-0fed6277096a'; ModuleVersion = '1.0' }
        @{ ModuleName = 'Dism'; Guid = '389c464d-8b8d-48e9-aafe-6d8a590d6798'; ModuleVersion = '1.0' }
        @{ ModuleName = 'International'; Guid = '561544e6-3a83-4d24-b140-78ad771eaf10'; ModuleVersion = '1.0' }
        @{ ModuleName = 'Microsoft.PowerShell.Archive'; Guid = 'eb74e8da-9ae2-482a-a648-e96550fb8733'; ModuleVersion = '1.0' }
        @{ ModuleName = 'Microsoft.PowerShell.Management'; Guid = 'eefcb906-b326-4e99-9f54-8b4bb6ef3c6d'; ModuleVersion = '1.0' }
        @{ ModuleName = 'Microsoft.PowerShell.Security'; Guid = 'a94c8c7e-9810-47c0-b8af-65089c13a35a'; ModuleVersion = '1.0' }
        @{ ModuleName = 'Microsoft.PowerShell.Utility'; Guid = '1da87e53-152b-403e-98dc-74d7b4d63d59'; ModuleVersion = '1.0' }
        @{ ModuleName = 'NetAdapter'; Guid = '1042b422-63a8-4016-a6d6-293e19e8f8a6'; ModuleVersion = '1.0' }
        @{ ModuleName = 'ScheduledTasks'; Guid = '5378ee8e-e349-49bb-83b9-f3d9c396c0a6'; ModuleVersion = '1.0' }
    )
    (& $Script:CommandTable.'Import-Module' -FullyQualifiedName $RequiredModules -Global -Force -PassThru -ErrorAction Stop).ExportedCommands.Values | & { process { $CommandTable.Add($_.Name, $_) } }
}

# Set required variables to ensure module functionality.
& $Script:CommandTable.'New-Variable' -Name ErrorActionPreference -Value ([System.Management.Automation.ActionPreference]::Stop) -Option Constant -Force
& $Script:CommandTable.'New-Variable' -Name InformationPreference -Value ([System.Management.Automation.ActionPreference]::Continue) -Option Constant -Force
& $Script:CommandTable.'New-Variable' -Name ProgressPreference -Value ([System.Management.Automation.ActionPreference]::SilentlyContinue) -Option Constant -Force

# Ensure module operates under the strictest of conditions.
& $Script:CommandTable.'Set-StrictMode' -Version 3

# Throw if any previous version of the unofficial PSADT module is found on the system.
if (& $Script:CommandTable.'Get-Module' -FullyQualifiedName @{ ModuleName = 'PSADT'; Guid = '41b2dd67-8447-4c66-b08a-f0bd0d5458b9'; ModuleVersion = '1.0' } -ListAvailable -Refresh)
{
    & $Script:CommandTable.'Write-Warning' -Message "This module should not be used while the unofficial v3 PSADT module is installed."
}

# Store build information pertaining to this module's state.
& $Script:CommandTable.'New-Variable' -Name Module -Option Constant -Force -Value ([ordered]@{
        Manifest = & $Script:CommandTable.'Import-LocalizedData' -BaseDirectory $PSScriptRoot -FileName 'PSAppDeployToolkit'
        Assemblies = (& $Script:CommandTable.'Get-ChildItem' -Path $PSScriptRoot\lib\PSADT*.dll).FullName
        Compiled = $MyInvocation.MyCommand.Name.Equals('PSAppDeployToolkit.psm1')
        Signed = (& $Script:CommandTable.'Get-AuthenticodeSignature' -LiteralPath $MyInvocation.MyCommand.Path).Status.Equals([System.Management.Automation.SignatureStatus]::Valid)
    }).AsReadOnly()

# Import our assemblies, factoring in whether they're on a network share or not.
$Module.Assemblies | & {
    begin
    {
        # Cache loaded assemblies to test whether they're already loaded.
        $domainAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()

        # Determine whether we're on a network location.
        $isNetworkLocation = [System.Uri]::new($PSScriptRoot).IsUnc -or ($PSScriptRoot -match '^([A-Za-z]:)\\' -and ((& $Script:CommandTable.'Get-CimInstance' -ClassName Win32_LogicalDisk -Filter "DeviceID='$($Matches[1])'").ProviderName -match '^\\\\'))
    }

    process
    {
        # Test whether the assembly is already loaded.
        if (($existingAssembly = $domainAssemblies | & { process { if ([System.IO.Path]::GetFileName($_.Location).Equals([System.IO.Path]::GetFileName($args[0]))) { return $_ } } } $_ | & $Script:CommandTable.'Select-Object' -First 1))
        {
            # Test the loaded assembly for SHA256 hash equality, returning early if the assembly is OK.
            if (!(& $Script:CommandTable.'Get-FileHash' -LiteralPath $existingAssembly.Location).Hash.Equals((& $Script:CommandTable.'Get-FileHash' -LiteralPath $_).Hash))
            {
                throw [System.Management.Automation.ErrorRecord]::new(
                    [System.InvalidOperationException]::new("A PSAppDeployToolkit assembly of a different file hash is already loaded. Please restart PowerShell and try again."),
                    'ConflictingModuleLoaded',
                    [System.Management.Automation.ErrorCategory]::InvalidOperation,
                    $existingAssembly
                )
            }
            return
        }

        # If we're on a compiled build, confirm the DLLs are signed before proceeding.
        if ($Module.Signed -and !($badFile = & $Script:CommandTable.'Get-AuthenticodeSignature' -LiteralPath $_).Status.Equals([System.Management.Automation.SignatureStatus]::Valid))
        {
            throw [System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new("The assembly [$_] has an invalid digital signature and cannot be loaded."),
                'ADTAssemblyFileSignatureError',
                [System.Management.Automation.ErrorCategory]::SecurityError,
                $badFile
            )
        }

        # If loading from an SMB path, load unsafely. This is OK because in signed (release) modules, we're validating the signature above.
        if ($isNetworkLocation)
        {
            [System.Reflection.Assembly]::UnsafeLoadFrom($_)
        }
        else
        {
            & $Script:CommandTable.'Add-Type' -LiteralPath $_
        }
    }
}

# Set the process as HiDPI so long as we're in a real console.
if ($Host.Name.Equals('ConsoleHost'))
{
    try
    {
        [PSADT.GUI.UiAutomation]::SetProcessDpiAwarenessForOSVersion()
    }
    catch
    {
        $null = $null
    }
}

# All WinForms-specific initialization code.
try
{
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
catch
{
    $null = $null
}

# Remove any previous functions that may have been defined.
if ($Module.Compiled)
{
    & $Script:CommandTable.'New-Variable' -Name FunctionPaths -Option Constant -Value ($MyInvocation.MyCommand.ScriptBlock.Ast.EndBlock.Statements | & { process { if ($_ -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return "Microsoft.PowerShell.Core\Function::$($_.Name)" } } })
    & $Script:CommandTable.'Remove-Item' -LiteralPath $FunctionPaths -Force -ErrorAction Ignore
}


#-----------------------------------------------------------------------------
#
# MARK: Close-ADTInstallationProgressClassic
#
#-----------------------------------------------------------------------------

function Close-ADTInstallationProgressClassic
{
    # Process the WPF window if it exists.
    if ($Script:Dialogs.Classic.ProgressWindow.SyncHash.ContainsKey('Window'))
    {
        if (!$Script:Dialogs.Classic.ProgressWindow.Invocation.IsCompleted)
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Closing the installation progress dialog.'
            $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window.Dispatcher.Invoke({ $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window.Close() }, [System.Windows.Threading.DispatcherPriority]::Send)
            while (!$Script:Dialogs.Classic.ProgressWindow.Invocation.IsCompleted) {}
        }
        $Script:Dialogs.Classic.ProgressWindow.SyncHash.Clear()
    }

    # End the PowerShell instance if it's invoked.
    if ($Script:Dialogs.Classic.ProgressWindow.Invocation)
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Closing the installation progress dialog's invocation."
        $null = $Script:Dialogs.Classic.ProgressWindow.PowerShell.EndInvoke($Script:Dialogs.Classic.ProgressWindow.Invocation)
        $Script:Dialogs.Classic.ProgressWindow.Invocation = $null
    }

    # Process the PowerShell window.
    if ($Script:Dialogs.Classic.ProgressWindow.PowerShell)
    {
        # Close down the runspace.
        if ($Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace -and $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.RunspaceStateInfo.State.Equals([System.Management.Automation.Runspaces.RunspaceState]::Opened))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Closing the installation progress dialog's runspace."
            $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.Close()
            $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.Dispose()
            $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace = $null
        }

        # Dispose of remaining PowerShell variables.
        $Script:Dialogs.Classic.ProgressWindow.PowerShell.Dispose()
        $Script:Dialogs.Classic.ProgressWindow.PowerShell = $null
    }

    # Reset the state bool.
    $Script:Dialogs.Classic.ProgressWindow.Running = $false
}


#-----------------------------------------------------------------------------
#
# MARK: Close-ADTInstallationProgressFluent
#
#-----------------------------------------------------------------------------

function Close-ADTInstallationProgressFluent
{
    # Hide the dialog and reset the state bool.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Closing the installation progress dialog.'
    [PSADT.UserInterface.UnifiedADTApplication]::CloseProgressDialog()
    $Script:Dialogs.Fluent.ProgressWindow.Running = $false
}


#-----------------------------------------------------------------------------
#
# MARK: Convert-RegistryKeyToHashtable
#
#-----------------------------------------------------------------------------

function Convert-RegistryKeyToHashtable
{
    begin
    {
        # Open collector to store all converted keys.
        $data = @{}
    }

    process
    {
        # Process potential subkeys first.
        $subdata = $_ | & $Script:CommandTable.'Get-ChildItem' | & $MyInvocation.MyCommand

        # Open a new subdata hashtable if we had no subkeys.
        if ($null -eq $subdata)
        {
            $subdata = @{}
        }

        # Process this item and store its values.
        $_ | & $Script:CommandTable.'Get-ItemProperty' | & {
            process
            {
                $_.PSObject.Properties | & {
                    process
                    {
                        if (($_.Name -notmatch '^PS((Parent)?Path|ChildName|Provider)$') -and ![System.String]::IsNullOrWhiteSpace((& $Script:CommandTable.'Out-String' -InputObject $_.Value)))
                        {
                            # Handle bools as string values.
                            if ($_.Value -match '^(True|False)$')
                            {
                                $subdata.Add($_.Name, [System.Boolean]::Parse($_.Value))
                            }
                            else
                            {
                                $subdata.Add($_.Name, $_.Value)
                            }
                        }
                    }
                }
            }
        }

        # Add the subdata to the sections if it's got a count.
        if ($subdata.Count)
        {
            $data.Add($_.PSPath -replace '^.+\\', $subdata)
        }
    }

    end
    {
        # If there's something in the collector, return it.
        if ($data.Count)
        {
            return $data
        }
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Disable-ADTWindowCloseButton
#
#-----------------------------------------------------------------------------

function Disable-ADTWindowCloseButton
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (($null -eq $_) -or $_.Equals([System.IntPtr]::Zero))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName WindowHandle -ProvidedValue $_ -ExceptionMessage 'The provided window handle is invalid.'))
                }
                return !!$_
            })]
        [System.IntPtr]$WindowHandle
    )

    $null = if (($menuHandle = [PSADT.LibraryInterfaces.User32]::GetSystemMenu($WindowHandle, $false)) -and ($menuHandle -ne [System.IntPtr]::Zero))
    {
        [PSADT.LibraryInterfaces.User32]::EnableMenuItem($menuHandle, 0xF060, 0x00000001)
        [PSADT.LibraryInterfaces.User32]::DestroyMenu($menuHandle)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Exit-ADTInvocation
#
#-----------------------------------------------------------------------------

function Exit-ADTInvocation
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [System.Nullable[System.Int32]]$ExitCode,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Force
    )

    # Attempt to close down any progress dialog here as an additional safety item.
    $progressOpen = if (& $Script:CommandTable.'Test-ADTInstallationProgressRunning')
    {
        try
        {
            & $Script:CommandTable.'Close-ADTInstallationProgress'
        }
        catch
        {
            $_
        }
    }

    # Flag the module as uninitialized upon last session closure.
    $Script:ADT.Initialized = $false

    # Return early if this function was called from the command line.
    if (($null -eq $ExitCode) -and !$Force)
    {
        return
    }

    # If a callback failed and we're in a proper console, forcibly exit the process.
    # The proper closure of a blocking dialog can stall a traditional exit indefinitely.
    if ($Force -or ($Host.Name.Equals('ConsoleHost') -and $progressOpen))
    {
        [System.Environment]::Exit($ExitCode)
    }
    exit $ExitCode
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTEdgeExtensions
#
#-----------------------------------------------------------------------------

function Get-ADTEdgeExtensions
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    param
    (
    )

    # Check if the ExtensionSettings registry key exists. If not, create it.
    if (!(& $Script:CommandTable.'Test-ADTRegistryValue' -Key Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge -Name ExtensionSettings))
    {
        & $Script:CommandTable.'Set-ADTRegistryKey' -Key Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge -Name ExtensionSettings -Value "" | & $Script:CommandTable.'Out-Null'
        return [pscustomobject]@{}
    }
    $extensionSettings = & $Script:CommandTable.'Get-ADTRegistryKey' -Key Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge -Name ExtensionSettings
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Configured extensions: [$($extensionSettings)]." -Severity 1
    return $extensionSettings | & $Script:CommandTable.'ConvertFrom-Json'
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTMountedWimFile
#
#-----------------------------------------------------------------------------

function Get-ADTMountedWimFile
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ImagePath', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Path', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    [OutputType([Microsoft.Dism.Commands.MountedImageInfoObject])]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'ImagePath')]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$ImagePath,

        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [System.IO.DirectoryInfo[]]$Path
    )

    # Get the caller's provided input via the ParameterSetName so we can filter on its name and value.
    $parameter = & $Script:CommandTable.'Get-Variable' -Name $PSCmdlet.ParameterSetName
    return (& $Script:CommandTable.'Get-WindowsImage' -Mounted | & { process { if ($parameter.Value.FullName.Contains($_.($parameter.Name))) { return $_ } } })
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTParentProcesses
#
#-----------------------------------------------------------------------------

function Get-ADTParentProcesses
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    param
    (
    )

    # Open object to store all parents for returning. This also avoids an infinite loop situation.
    $parents = [System.Collections.Generic.List[Microsoft.Management.Infrastructure.CimInstance]]::new()

    # Get all processes from the system. WMI consistently gives us the parent on PowerShell 5.x and Core targets.
    $processes = & $Script:CommandTable.'Get-CimInstance' -ClassName Win32_Process
    $process = $processes | & { process { if ($_.ProcessId -eq $PID) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1

    # Get all parents for the currently stored process.
    while ($process = $processes | & { process { if ($_.ProcessId -eq $process.ParentProcessId) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1)
    {
        if ($parents.Contains($process))
        {
            break
        }
        $parents.Add($process)
    }

    # Return all parents to the caller.
    return $parents
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTProcessHandles
#
#-----------------------------------------------------------------------------

function Get-ADTProcessHandles
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    param
    (
    )

    # Get CSV data from the binary and confirm success.
    $exeHandle = "$Script:PSScriptRoot\bin\$([PSADT.OperatingSystem.OSHelper]::GetArchitecture())\handle\handle.exe"
    $exeHandleResults = & $exeHandle -accepteula -nobanner -v 2>&1
    if ($Global:LASTEXITCODE -ne 0)
    {
        $naerParams = @{
            Exception = [System.Runtime.InteropServices.ExternalException]::new("The call to [$exeHandle] failed with exit code [$Global:LASTEXITCODE].", $Global:LASTEXITCODE)
            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
            ErrorId = 'HandleExecutableFailure'
            TargetObject = $exeHandleResults
            RecommendedAction = "Please review the result in this error's TargetObject property and try again."
        }
        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
    }

    # Convert CSV data to objects and re-process to remove non-word characters before returning data to the caller.
    if (($handles = $exeHandleResults | & $Script:CommandTable.'ConvertFrom-Csv'))
    {
        return $handles | & $Script:CommandTable.'Select-Object' -Property ($handles[0].PSObject.Properties.Name | & {
                process
                {
                    @{ Label = $_ -replace '[^\w]'; Expression = [scriptblock]::Create("`$_.'$_'.Trim()") }
                }
            })
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTRunningProcesses
#
#-----------------------------------------------------------------------------

function Get-ADTRunningProcesses
{
    <#

    .SYNOPSIS
    Gets the processes that are running from a custom list of process objects and also adds a property called ProcessDescription.

    .DESCRIPTION
    Gets the processes that are running from a custom list of process objects and also adds a property called ProcessDescription.

    .PARAMETER ProcessObjects
    Custom object containing the process objects to search for.

    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    System.Diagnostics.Process. Returns one or more process objects representing each running process found.

    .EXAMPLE
    Get-ADTRunningProcesses -ProcessObjects $processObjects

    .NOTES
    This is an internal script function and should typically not be called directly.

    .NOTES
    An active ADT session is NOT required to use this function.

    .LINK
    https://psappdeploytoolkit.com

    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([System.Diagnostics.Process])]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowNull()][AllowEmptyCollection()]
        [PSADT.Types.ProcessObject[]]$ProcessObjects
    )

    # Return early if we've received no input.
    if ($null -eq $ProcessObjects)
    {
        return
    }

    # Get all running processes and append properties.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Checking for running applications: [$($ProcessObjects.Name -join ',')]"
    $runningProcesses = & $Script:CommandTable.'Get-Process' -Name $ProcessObjects.Name -ErrorAction Ignore | & {
        process
        {
            if (!$_.HasExited)
            {
                return $_ | & $Script:CommandTable.'Add-Member' -MemberType NoteProperty -Name ProcessDescription -Force -PassThru -Value $(
                    if (![System.String]::IsNullOrWhiteSpace(($objDescription = $ProcessObjects | & $Script:CommandTable.'Where-Object' -Property Name -EQ -Value $_.ProcessName | & $Script:CommandTable.'Select-Object' -ExpandProperty Description -ErrorAction Ignore)))
                    {
                        # The description of the process provided with the object.
                        $objDescription
                    }
                    elseif ($_.Description)
                    {
                        # If the process already has a description field specified, then use it.
                        $_.Description
                    }
                    else
                    {
                        # Fall back on the process name if no description is provided by the process or as a parameter to the function.
                        $_.ProcessName
                    }
                )
            }
        }
    }

    # Return output if there's any.
    if ($runningProcesses)
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "The following processes are running: [$(($runningProcesses.ProcessName | & $Script:CommandTable.'Select-Object' -Unique) -join ',')]."
        return ($runningProcesses | & $Script:CommandTable.'Sort-Object' -Property ProcessDescription)
    }
    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Specified applications are not running.'
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTSCCMClientVersion
#
#-----------------------------------------------------------------------------

function Get-ADTSCCMClientVersion
{
    # Make sure SCCM client is installed and running.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Checking to see if SCCM Client service [ccmexec] is installed and running.'
    if (!(& $Script:CommandTable.'Test-ADTServiceExists' -Name ccmexec))
    {
        $naerParams = @{
            Exception = [System.ApplicationException]::new('SCCM Client Service [ccmexec] does not exist. The SCCM Client may not be installed.')
            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
            ErrorId = 'CcmExecServiceMissing'
            RecommendedAction = "Please check the availability of this service and try again."
        }
        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
    }
    if (($svc = & $Script:CommandTable.'Get-Service' -Name ccmexec).Status -ne 'Running')
    {
        $naerParams = @{
            Exception = [System.ApplicationException]::new("SCCM Client Service [ccmexec] exists but it is not in a 'Running' state.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
            ErrorId = 'CcmExecServiceNotRunning'
            TargetObject = $svc
            RecommendedAction = "Please check the status of this service and try again."
        }
        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
    }

    # Determine the SCCM Client Version.
    try
    {
        [System.Version]$SCCMClientVersion = & $Script:CommandTable.'Get-CimInstance' -Namespace ROOT\CCM -ClassName CCM_InstalledComponent | & { process { if ($_.Name -eq 'SmsClient') { $_.Version } } }
    }
    catch
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to query the system for the SCCM client version number.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 2
        throw
    }
    if (!$SCCMClientVersion)
    {
        $naerParams = @{
            Exception = [System.Data.NoNullAllowedException]::new('The query for the SmsClient version returned a null result.')
            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
            ErrorId = 'CcmExecVersionNullOrEmpty'
            RecommendedAction = "Please check the installed version and try again."
        }
        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
    }
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Installed SCCM Client Version Number [$SCCMClientVersion]."
    return $SCCMClientVersion
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTStringLanguage
#
#-----------------------------------------------------------------------------

function Get-ADTStringLanguage
{
    if (![System.String]::IsNullOrWhiteSpace(($adtConfig = & $Script:CommandTable.'Get-ADTConfig').UI.LanguageOverride))
    {
        # The caller has specified a specific language.
        return $adtConfig.UI.LanguageOverride
    }
    else
    {
        # Fall back to PowerShell's.
        return [System.Threading.Thread]::CurrentThread.CurrentUICulture
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Import-ADTConfig
#
#-----------------------------------------------------------------------------

function Import-ADTConfig
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName BaseDirectory -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName BaseDirectory -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String[]]$BaseDirectory
    )

    # Internal filter to process asset file paths.
    filter Update-ADTAssetFilePath
    {
        # Go recursive if we've received a hashtable, otherwise just update the values.
        foreach ($asset in $($_.GetEnumerator()))
        {
            # Re-process if this is a hashtable.
            if ($asset.Value -is [System.Collections.Hashtable])
            {
                $asset.Value | & $MyInvocation.MyCommand; continue
            }

            # Skip if the path is fully qualified.
            if ([System.IO.Path]::IsPathRooted($asset.Value))
            {
                continue
            }

            # Get the asset's full path based on the supplied BaseDirectory.
            # Fall back to the module's path if the asset is unable to be found.
            $assetPath = foreach ($directory in $($BaseDirectory[($BaseDirectory.Count - 1)..(0)]; $Script:ADT.Directories.Defaults.Config))
            {
                if (($assetPath = & $Script:CommandTable.'Get-Item' -LiteralPath "$directory\$($_.($asset.Key))" -ErrorAction Ignore))
                {
                    $assetPath.FullName
                    break
                }
            }

            # Throw if we found no asset.
            if (!$assetPath)
            {
                $naerParams = @{
                    Exception = [System.IO.FileNotFoundException]::new("Failed to resolve the asset [$($asset.Key)] to a valid file path.", $_.($asset.Key))
                    Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                    ErrorId = 'DialogAssetNotFound'
                    TargetObject = $_.($asset.Key)
                    RecommendedAction = "Ensure the file exists and try again."
                }
                $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
            }
            $_.($asset.Key) = $assetPath
        }
    }

    # Internal filter to expand variables.
    filter Expand-ADTVariablesInConfig
    {
        # Go recursive if we've received a hashtable, otherwise just update the values.
        foreach ($section in $($_.GetEnumerator()))
        {
            if ($section.Value -is [System.Collections.Hashtable])
            {
                $section.Value | & $MyInvocation.MyCommand
            }
            elseif ($section.Value -is [System.String])
            {
                $_.($section.Key) = $ExecutionContext.InvokeCommand.ExpandString($section.Value)
            }
        }
    }

    # Import the config from disk.
    $config = & $Script:CommandTable.'Import-ADTModuleDataFile' @PSBoundParameters -FileName config.psd1

    # Place restrictions on non-ConsoleHost targets.
    if ($Host.Name.Equals('Windows PowerShell ISE Host'))
    {
        $config.UI.DialogStyle = 'Classic'
    }

    # Confirm the specified dialog type is valid.
    if (($config.UI.DialogStyle -ne 'Classic') -and (& $Script:CommandTable.'Test-ADTNonNativeCaller'))
    {
        $config.UI.DialogStyle = if ($config.UI.ContainsKey('DialogStyleCompatMode'))
        {
            $config.UI.DialogStyleCompatMode
        }
        else
        {
            'Classic'
        }
    }
    if (!$Script:Dialogs.Contains($config.UI.DialogStyle))
    {
        $naerParams = @{
            Exception = [System.NotSupportedException]::new("The specified dialog style [$($config.UI.DialogStyle)] is not supported. Valid styles are ['$($Script:Dialogs.Keys -join "', '")'].")
            Category = [System.Management.Automation.ErrorCategory]::InvalidData
            ErrorId = 'DialogStyleNotSupported'
            TargetObject = $config
            RecommendedAction = "Please review the supplied configuration file and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }

    # Expand out environment variables and asset file paths.
    ($adtEnv = & $Script:CommandTable.'Get-ADTEnvironmentTable').GetEnumerator() | & { process { & $Script:CommandTable.'New-Variable' -Name $_.Name -Value $_.Value -Option Constant } end { $config | Expand-ADTVariablesInConfig } }
    $config.Assets | Update-ADTAssetFilePath

    # Process the classic assets by grabbing the bytes of each image asset, storing them into a memory stream, then as an image for WinForms to use.
    $Script:Dialogs.Classic.Assets.Logo = [System.Drawing.Image]::FromStream([System.IO.MemoryStream]::new([System.IO.File]::ReadAllBytes($config.Assets.Logo)))
    $Script:Dialogs.Classic.Assets.Icon = [PSADT.Shared.Utility]::ConvertImageToIcon($Script:Dialogs.Classic.Assets.Logo)
    $Script:Dialogs.Classic.Assets.Banner = [System.Drawing.Image]::FromStream([System.IO.MemoryStream]::new([System.IO.File]::ReadAllBytes($config.Assets.Banner)))
    $Script:Dialogs.Classic.BannerHeight = [System.Math]::Ceiling($Script:Dialogs.Classic.Width * ($Script:Dialogs.Classic.Assets.Banner.Height / $Script:Dialogs.Classic.Assets.Banner.Width))

    # Set the app's AUMID so it doesn't just say "Windows PowerShell".
    if ($config.UI.BalloonNotifications -and ![PSADT.LibraryInterfaces.Shell32]::SetCurrentProcessExplicitAppUserModelID($config.UI.BalloonTitle))
    {
        $regKey = "$(if ($adtEnv.IsAdmin) { 'HKEY_CLASSES_ROOT' } else { 'HKEY_CURRENT_USER\Software\Classes' })\AppUserModelId\$($config.UI.BalloonTitle)"
        [Microsoft.Win32.Registry]::SetValue($regKey, 'DisplayName', $config.UI.BalloonTitle, [Microsoft.Win32.RegistryValueKind]::String)
        [Microsoft.Win32.Registry]::SetValue($regKey, 'IconUri', $config.Assets.Logo, [Microsoft.Win32.RegistryValueKind]::ExpandString)
    }

    # Change paths to user accessible ones if user isn't an admin.
    if (!$adtEnv.IsAdmin)
    {
        if ($config.Toolkit.TempPathNoAdminRights)
        {
            $config.Toolkit.TempPath = $config.Toolkit.TempPathNoAdminRights
        }
        if ($config.Toolkit.RegPathNoAdminRights)
        {
            $config.Toolkit.RegPath = $config.Toolkit.RegPathNoAdminRights
        }
        if ($config.Toolkit.LogPathNoAdminRights)
        {
            $config.Toolkit.LogPath = $config.Toolkit.LogPathNoAdminRights
        }
        if ($config.MSI.LogPathNoAdminRights)
        {
            $config.MSI.LogPath = $config.MSI.LogPathNoAdminRights
        }
    }

    # Append the toolkit's name onto the temporary path.
    $config.Toolkit.TempPath = [System.IO.Path]::Combine($config.Toolkit.TempPath, $adtEnv.appDeployToolkitName)

    # Finally, return the config for usage within module.
    return $config
}


#-----------------------------------------------------------------------------
#
# MARK: Import-ADTModuleDataFile
#
#-----------------------------------------------------------------------------

function Import-ADTModuleDataFile
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName BaseDirectory -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName BaseDirectory -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String[]]$BaseDirectory,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FileName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$UICulture,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IgnorePolicy
    )

    # Internal function to process the imported data.
    function Update-ADTImportedDataValues
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.Hashtable]$DataFile,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.Collections.Hashtable]$NewData
        )

        # Process the provided default data so we can add missing data to the data file.
        foreach ($section in $NewData.GetEnumerator())
        {
            # Recursively process hashtables, otherwise just update the value.
            if ($section.Value -is [System.Collections.Hashtable])
            {
                if (!$DataFile.ContainsKey($section.Key) -or ($DataFile.($section.Key) -isnot [System.Collections.Hashtable]))
                {
                    $DataFile.($section.Key) = @{}
                }
                & $MyInvocation.MyCommand -DataFile $DataFile.($section.Key) -NewData $section.Value
            }
            elseif (!$DataFile.ContainsKey($section.Key) -or ![System.String]::IsNullOrWhiteSpace((& $Script:CommandTable.'Out-String' -InputObject $section.Value)))
            {
                $DataFile.($section.Key) = $section.Value
            }
        }
    }

    # Establish directory paths for the specified input.
    $moduleDirectory = $Script:ADT.Directories.Defaults.([System.IO.Path]::GetFileNameWithoutExtension($FileName))
    $callerDirectory = $BaseDirectory

    # If we're running a release module, ensure the psd1 files haven't been tampered with.
    if (($badFiles = & $Script:CommandTable.'Test-ADTReleaseBuildFileValidity' -LiteralPath $moduleDirectory))
    {
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("The module's default $FileName file has been modified from its released state.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidData
            ErrorId = 'ADTDataFileSignatureError'
            TargetObject = $badFiles
            RecommendedAction = "Please re-download $($MyInvocation.MyCommand.Module.Name) and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }

    # Import the default data first and foremost.
    $null = $PSBoundParameters.Remove('IgnorePolicy')
    $PSBoundParameters.BaseDirectory = $moduleDirectory
    $importedData = & $Script:CommandTable.'Import-LocalizedData' @PSBoundParameters

    # Validate we imported something from our default location.
    if (!$importedData.Count)
    {
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("The importation of the module's default $FileName file returned a null or empty result.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
            ErrorId = 'ADTDataFileImportFailure'
            TargetObject = [System.IO.Path]::Combine($PSBoundParameters.BaseDirectory, $FileName)
            RecommendedAction = "Please ensure that this module is not corrupt or missing files, then try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }

    # Super-impose the caller's data if it's different from default.
    if (!$callerDirectory.Equals($moduleDirectory))
    {
        foreach ($directory in $callerDirectory)
        {
            $PSBoundParameters.BaseDirectory = $directory
            Update-ADTImportedDataValues -DataFile $importedData -NewData (& $Script:CommandTable.'Import-LocalizedData' @PSBoundParameters)
        }
    }

    # Super-impose registry values if they exist.
    if (!$IgnorePolicy -and ($policySettings = & $Script:CommandTable.'Get-ChildItem' -LiteralPath "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\PSAppDeployToolkit\$([System.IO.Path]::GetFileNameWithoutExtension($FileName))" -ErrorAction Ignore | & $Script:CommandTable.'Convert-RegistryKeyToHashtable'))
    {
        Update-ADTImportedDataValues -DataFile $importedData -NewData $policySettings
    }

    # Return the built out data to the caller.
    return $importedData
}


#-----------------------------------------------------------------------------
#
# MARK: Initialize-ADTModuleIfUnitialized
#
#-----------------------------------------------------------------------------

function Initialize-ADTModuleIfUnitialized
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )

    # Initialize the module if there's no session and it hasn't been previously initialized.
    if (!($adtSession = if (& $Script:CommandTable.'Test-ADTSessionActive') { & $Script:CommandTable.'Get-ADTSession' }) -and !(& $Script:CommandTable.'Test-ADTModuleInitialized'))
    {
        try
        {
            & $Script:CommandTable.'Initialize-ADTModule'
        }
        catch
        {
            $Cmdlet.ThrowTerminatingError($_)
        }
    }

    # Return the current session if we happened to get one.
    if ($adtSession)
    {
        return $adtSession
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTServiceAndDependencyOperation
#
#-----------------------------------------------------------------------------

function Invoke-ADTServiceAndDependencyOperation
{
    <#

    .SYNOPSIS
    Process Windows service and its dependencies.

    .DESCRIPTION
    Process Windows service and its dependencies.

    .PARAMETER Service
    Specify the name of the service.

    .PARAMETER SkipDependentServices
    Choose to skip checking for dependent services. Default is: $false.

    .PARAMETER PendingStatusWait
    The amount of time to wait for a service to get out of a pending state before continuing. Default is 60 seconds.

    .PARAMETER PassThru
    Return the System.ServiceProcess.ServiceController service object.

    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    System.ServiceProcess.ServiceController. Returns the service object.

    .EXAMPLE
    Invoke-ADTServiceAndDependencyOperation -Service wuauserv -Operation Start

    .EXAMPLE
    Invoke-ADTServiceAndDependencyOperation -Service wuauserv -Operation Stop

    .LINK
    https://psappdeploytoolkit.com

    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!$_.Name)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Service -ProvidedValue $_ -ExceptionMessage 'The specified service does not exist.'))
                }
                return !!$_
            })]
        [System.ServiceProcess.ServiceController]$Service,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Start', 'Stop')]
        [System.String]$Operation,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipDependentServices,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$PendingStatusWait = [System.TimeSpan]::FromSeconds(60),

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    # Internal worker function.
    function Invoke-ADTDependentServiceOperation
    {
        # Discover all dependent services.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Discovering all dependent service(s) for service [$Service] which are not '$(($status = if ($Operation -eq 'Start') {'Running'} else {'Stopped'}))'."
        if (!($dependentServices = & $Script:CommandTable.'Get-Service' -Name $Service.ServiceName -DependentServices | & { process { if ($_.Status -ne $status) { return $_ } } }))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Dependent service(s) were not discovered for service [$Service]."
            return
        }

        # Action each found dependent service.
        foreach ($dependent in $dependentServices)
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(('Starting', 'Stopping')[$Operation -eq 'Start']) dependent service [$($dependent.ServiceName)] with display name [$($dependent.DisplayName)] and a status of [$($dependent.Status)]."
            try
            {
                $dependent | & "$($Operation)-Service" -Force -WarningAction Ignore
            }
            catch
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to $($Operation.ToLower()) dependent service [$($dependent.ServiceName)] with display name [$($dependent.DisplayName)] and a status of [$($dependent.Status)]. Continue..." -Severity 2
            }
        }
    }

    # Wait up to 60 seconds if service is in a pending state.
    if (($desiredStatus = @{ ContinuePending = 'Running'; PausePending = 'Paused'; StartPending = 'Running'; StopPending = 'Stopped' }[$Service.Status]))
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Waiting for up to [$($PendingStatusWait.TotalSeconds)] seconds to allow service pending status [$($Service.Status)] to reach desired status [$([System.ServiceProcess.ServiceControllerStatus]$desiredStatus)]."
        $Service.WaitForStatus($desiredStatus, $PendingStatusWait)
        $Service.Refresh()
    }

    # Discover if the service is currently running.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Service [$($Service.ServiceName)] with display name [$($Service.DisplayName)] has a status of [$($Service.Status)]."
    if (($Operation -eq 'Stop') -and ($Service.Status -ne 'Stopped'))
    {
        # Process all dependent services.
        if (!$SkipDependentServices)
        {
            Invoke-ADTDependentServiceOperation
        }

        # Stop the parent service.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Stopping parent service [$($Service.ServiceName)] with display name [$($Service.DisplayName)]."
        $Service = $Service | & $Script:CommandTable.'Stop-Service' -PassThru -WarningAction Ignore -Force
    }
    elseif (($Operation -eq 'Start') -and ($Service.Status -ne 'Running'))
    {
        # Start the parent service.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Starting parent service [$($Service.ServiceName)] with display name [$($Service.DisplayName)]."
        $Service = $Service | & $Script:CommandTable.'Start-Service' -PassThru -WarningAction Ignore

        # Process all dependent services.
        if (!$SkipDependentServices)
        {
            Invoke-ADTDependentServiceOperation
        }
    }

    # Return the service object if option selected.
    if ($PassThru)
    {
        return $Service
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTSessionCallbackOperation
#
#-----------------------------------------------------------------------------

function Invoke-ADTSessionCallbackOperation
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Action', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('Starting', 'Opening', 'Closing', 'Finishing')]
        [System.String]$Type,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Add', 'Remove')]
        [System.String]$Action,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Cache the global callbacks and perform any required action.
    $callbacks = $Script:ADT.Callbacks.$Type
    $null = $Callback | & { process { if ($Action.Equals('Remove') -or !$callbacks.Contains($_)) { $callbacks.$Action($_) } } }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTSubstOperation
#
#-----------------------------------------------------------------------------

function Invoke-ADTSubstOperation
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Create')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Delete')]
        [ValidateScript({
                if ($_ -notmatch '^[A-Z]:$')
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Drive -ProvidedValue $_ -ExceptionMessage 'The specified drive is not valid. Please specify a drive in the following format: [A:, B:, etc].'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$Drive,

        [Parameter(Mandatory = $true, ParameterSetName = 'Create')]
        [ValidateScript({
                if ($null -eq $_)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified input is null.'))
                }
                if (!$_.Exists)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified image path cannot be found.'))
                }
                if ([System.Uri]::new($_).IsUnc)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified image path cannot be a network share.'))
                }
                return !!$_
            })]
        [System.IO.DirectoryInfo]$Path,

        [Parameter(Mandatory = $true, ParameterSetName = 'Delete')]
        [System.Management.Automation.SwitchParameter]$Delete
    )

    # Perform the subst operation. An exit code of 0 is considered successful.
    $substPath = "$([System.Environment]::SystemDirectory)\subst.exe"
    $substResult = if ($Path)
    {
        # Throw if the specified drive letter is in use.
        if ((& $Script:CommandTable.'Get-PSDrive' -PSProvider FileSystem).Name -contains $Drive.Substring(0, 1))
        {
            $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Drive -ProvidedValue $Drive -ExceptionMessage 'The specified drive is currently in use. Please try again with an unused drive letter.'))
        }
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(($msg = "Creating substitution drive [$Drive] for [$Path]"))."
        & $substPath $Drive $Path.FullName
    }
    elseif ($Delete)
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(($msg = "Deleting substitution drive [$Drive]"))."
        & $substPath $Drive /D
    }
    else
    {
        # If we're here, the caller probably did something silly like -Delete:$false.
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("Unable to determine the required mode of operation.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
            ErrorId = 'SubstModeIndeterminate'
            TargetObject = $PSBoundParameters
            RecommendedAction = "Please review the result in this error's TargetObject property and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }
    if ($Global:LASTEXITCODE.Equals(0))
    {
        return
    }

    # If we're here, we had a bad exit code.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message ($msg = "$msg failed with exit code [$Global:LASTEXITCODE]: $substResult") -Severity 3
    $naerParams = @{
        Exception = [System.Runtime.InteropServices.ExternalException]::new($msg, $Global:LASTEXITCODE)
        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
        ErrorId = 'SubstUtilityFailure'
        TargetObject = $substResult
        RecommendedAction = "Please review the result in this error's TargetObject property and try again."
    }
    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTTerminalServerModeChange
#
#-----------------------------------------------------------------------------

function Invoke-ADTTerminalServerModeChange
{
    <#

    .SYNOPSIS
    Changes the mode for Remote Desktop Session Host/Citrix servers.

    .DESCRIPTION
    Changes the mode for Remote Desktop Session Host/Citrix servers.

    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    None. This function does not return any objects.

    .LINK
    https://psappdeploytoolkit.com

    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('Install', 'Execute')]
        [System.String]$Mode
    )

    # Change the terminal server mode. An exit code of 1 is considered successful.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(($msg = "Changing terminal server into user $($Mode.ToLower()) mode"))."
    $terminalServerResult = & "$([System.Environment]::SystemDirectory)\change.exe" User /$Mode 2>&1
    if ($Global:LASTEXITCODE.Equals(1))
    {
        return
    }

    # If we're here, we had a bad exit code.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message ($msg = "$msg failed with exit code [$Global:LASTEXITCODE]: $terminalServerResult") -Severity 3
    $naerParams = @{
        Exception = [System.Runtime.InteropServices.ExternalException]::new($msg, $Global:LASTEXITCODE)
        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
        ErrorId = 'RdsChangeUtilityFailure'
        TargetObject = $terminalServerResult
        RecommendedAction = "Please review the result in this error's TargetObject property and try again."
    }
    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTEnvironmentTable
#
#-----------------------------------------------------------------------------

function New-ADTEnvironmentTable
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This function does not change system state.")]
    param
    (
    )

    # Perform initial setup.
    $variables = [ordered]@{}

    ## Variables: Toolkit Info
    $variables.Add('appDeployToolkitName', $MyInvocation.MyCommand.Module.Name)
    $variables.Add('appDeployToolkitPath', $MyInvocation.MyCommand.Module.ModuleBase)
    $variables.Add('appDeployMainScriptVersion', $MyInvocation.MyCommand.Module.Version)

    ## Variables: Culture
    $variables.Add('culture', $Host.CurrentCulture)
    $variables.Add('uiculture', $Host.CurrentUICulture)
    $variables.Add('currentLanguage', $variables.culture.TwoLetterISOLanguageName.ToUpper())
    $variables.Add('currentUILanguage', $variables.uiculture.TwoLetterISOLanguageName.ToUpper())

    ## Variables: Environment Variables
    $variables.Add('envHost', $Host)
    $variables.Add('envHostVersion', [System.Version]$Host.Version)
    $variables.Add('envHostVersionSemantic', $(if ($Host.Version.PSObject.Properties.Name -match '^PSSemVer') { [System.Management.Automation.SemanticVersion]$Host.Version }))
    $variables.Add('envHostVersionMajor', $variables.envHostVersion.Major)
    $variables.Add('envHostVersionMinor', $variables.envHostVersion.Minor)
    $variables.Add('envHostVersionBuild', $(if ($variables.envHostVersion.Build -ge 0) { $variables.envHostVersion.Build }))
    $variables.Add('envHostVersionRevision', $(if ($variables.envHostVersion.Revision -ge 0) { $variables.envHostVersion.Revision }))
    $variables.Add('envHostVersionPreReleaseLabel', $(if ($variables.envHostVersionSemantic -and $variables.envHostVersionSemantic.PreReleaseLabel) { $variables.envHostVersionSemantic.PreReleaseLabel }))
    $variables.Add('envHostVersionBuildLabel', $(if ($variables.envHostVersionSemantic -and $variables.envHostVersionSemantic.BuildLabel) { $variables.envHostVersionSemantic.BuildLabel }))
    $variables.Add('envAllUsersProfile', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData))
    $variables.Add('envAppData', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData))
    $variables.Add('envArchitecture', [System.Environment]::GetEnvironmentVariable('PROCESSOR_ARCHITECTURE'))
    $variables.Add('envCommonDesktop', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonDesktopDirectory))
    $variables.Add('envCommonDocuments', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonDocuments))
    $variables.Add('envCommonStartMenuPrograms', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonPrograms))
    $variables.Add('envCommonStartMenu', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonStartMenu))
    $variables.Add('envCommonStartUp', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonStartup))
    $variables.Add('envCommonTemplates', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonTemplates))
    $variables.Add('envHomeDrive', [System.Environment]::GetEnvironmentVariable('HOMEDRIVE'))
    $variables.Add('envHomePath', [System.Environment]::GetEnvironmentVariable('HOMEPATH'))
    $variables.Add('envHomeShare', [System.Environment]::GetEnvironmentVariable('HOMESHARE'))
    $variables.Add('envLocalAppData', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData))
    $variables.Add('envLogicalDrives', [System.Environment]::GetLogicalDrives())
    $variables.Add('envProgramData', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData))
    $variables.Add('envPublic', [System.Environment]::GetEnvironmentVariable('PUBLIC'))
    $variables.Add('envSystemDrive', [System.IO.Path]::GetPathRoot([System.Environment]::SystemDirectory).TrimEnd('\'))
    $variables.Add('envSystemRoot', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows))
    $variables.Add('envTemp', [System.IO.Path]::GetTempPath())
    $variables.Add('envUserCookies', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Cookies))
    $variables.Add('envUserDesktop', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::DesktopDirectory))
    $variables.Add('envUserFavorites', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Favorites))
    $variables.Add('envUserInternetCache', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::InternetCache))
    $variables.Add('envUserInternetHistory', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::History))
    $variables.Add('envUserMyDocuments', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments))
    $variables.Add('envUserName', [System.Environment]::UserName)
    $variables.Add('envUserPictures', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyPictures))
    $variables.Add('envUserProfile', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))
    $variables.Add('envUserSendTo', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::SendTo))
    $variables.Add('envUserStartMenu', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::StartMenu))
    $variables.Add('envUserStartMenuPrograms', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Programs))
    $variables.Add('envUserStartUp', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::StartUp))
    $variables.Add('envUserTemplates', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Templates))
    $variables.Add('envSystem32Directory', [System.Environment]::SystemDirectory)
    $variables.Add('envWinDir', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows))

    ## Variables: Running in SCCM Task Sequence.
    $variables.Add('RunningTaskSequence', !![System.Type]::GetTypeFromProgID('Microsoft.SMS.TSEnvironment'))

    ## Variables: Domain Membership
    $w32cs = & $Script:CommandTable.'Get-CimInstance' -ClassName Win32_ComputerSystem -Verbose:$false
    $w32csd = $w32cs.Domain | & { process { if ($_) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1
    $variables.Add('IsMachinePartOfDomain', $w32cs.PartOfDomain)
    $variables.Add('envMachineWorkgroup', $null)
    $variables.Add('envMachineADDomain', $null)
    $variables.Add('envLogonServer', $null)
    $variables.Add('MachineDomainController', $null)
    $variables.Add('envMachineDNSDomain', ([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName | & { process { if ($_) { return $_.ToLower() } } } | & $Script:CommandTable.'Select-Object' -First 1))
    $variables.Add('envUserDNSDomain', ([System.Environment]::GetEnvironmentVariable('USERDNSDOMAIN') | & { process { if ($_) { return $_.ToLower() } } } | & $Script:CommandTable.'Select-Object' -First 1))
    $variables.Add('envUserDomain', $(if ([System.Environment]::UserDomainName) { [System.Environment]::UserDomainName.ToUpper() }))
    $variables.Add('envComputerName', $w32cs.DNSHostName.ToUpper())
    $variables.Add('envComputerNameFQDN', $variables.envComputerName)
    if ($variables.IsMachinePartOfDomain)
    {
        $variables.envMachineADDomain = $w32csd.ToLower()
        $variables.envComputerNameFQDN = try
        {
            [System.Net.Dns]::GetHostEntry('localhost').HostName
        }
        catch
        {
            # Function GetHostEntry failed, but we can construct the FQDN in another way
            $variables.envComputerNameFQDN + '.' + $variables.envMachineADDomain
        }

        # Set the logon server and remove backslashes at the beginning.
        $variables.envLogonServer = $(try
            {
                [System.Environment]::GetEnvironmentVariable('LOGONSERVER') | & { process { if ($_ -and !$_.Contains('\\MicrosoftAccount')) { [System.Net.Dns]::GetHostEntry($_.TrimStart('\')).HostName } } }
            }
            catch
            {
                # If running in system context or if GetHostEntry fails, fall back on the logonserver value stored in the registry
                & $Script:CommandTable.'Get-ItemProperty' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History' -ErrorAction Ignore | & $Script:CommandTable.'Select-Object' -ExpandProperty DCName -ErrorAction Ignore
            })
        while ($variables.envLogonServer -and $variables.envLogonServer.StartsWith('\'))
        {
            $variables.envLogonServer = $variables.envLogonServer.Substring(1)
        }

        try
        {
            $variables.MachineDomainController = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController().Name
        }
        catch
        {
            $null = $null
        }
    }
    else
    {
        $variables.envMachineWorkgroup = $w32csd.ToUpper()
    }

    # Get the OS Architecture.
    $variables.Add('Is64Bit', [System.Environment]::Is64BitOperatingSystem)
    $variables.Add('envOSArchitecture', [PSADT.OperatingSystem.OSHelper]::GetArchitecture())

    ## Variables: Current Process Architecture
    $variables.Add('Is64BitProcess', [System.Environment]::Is64BitProcess)
    $variables.Add('psArchitecture', (& $Script:CommandTable.'Get-ADTPEFileArchitecture' -FilePath ([System.Diagnostics.Process]::GetCurrentProcess().Path)))

    ## Variables: Get normalized paths that vary depending on process bitness.
    if ($variables.Is64Bit)
    {
        if ($variables.Is64BitProcess)
        {
            $variables.Add('envProgramFiles', [System.Environment]::GetFolderPath('ProgramFiles'))
            $variables.Add('envCommonProgramFiles', [System.Environment]::GetFolderPath('CommonProgramFiles'))
            $variables.Add('envSysNativeDirectory', [System.Environment]::SystemDirectory)
            $variables.Add('envSYSWOW64Directory', [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::SystemX86))
        }
        else
        {
            $variables.Add('envProgramFiles', [System.Environment]::GetEnvironmentVariable('ProgramW6432'))
            $variables.Add('envCommonProgramFiles', [System.Environment]::GetEnvironmentVariable('CommonProgramW6432'))
            $variables.Add('envSysNativeDirectory', [System.IO.Path]::Combine([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows), 'sysnative'))
            $variables.Add('envSYSWOW64Directory', [System.Environment]::SystemDirectory)
        }
        $variables.Add('envProgramFilesX86', [System.Environment]::GetFolderPath('ProgramFilesX86'))
        $variables.Add('envCommonProgramFilesX86', [System.Environment]::GetFolderPath('CommonProgramFilesX86'))
    }
    else
    {
        $variables.Add('envProgramFiles', [System.Environment]::GetFolderPath('ProgramFiles'))
        $variables.Add('envProgramFilesX86', $null)
        $variables.Add('envCommonProgramFiles', [System.Environment]::GetFolderPath('CommonProgramFiles'))
        $variables.Add('envCommonProgramFilesX86', $null)
        $variables.Add('envSysNativeDirectory', [System.Environment]::SystemDirectory)
        $variables.Add('envSYSWOW64Directory', $null)
    }

    ## Variables: Operating System
    $variables.Add('envOS', (& $Script:CommandTable.'Get-CimInstance' -ClassName Win32_OperatingSystem -Verbose:$false))
    $variables.Add('envOSName', $variables.envOS.Caption.Trim())
    $variables.Add('envOSServicePack', $variables.envOS.CSDVersion)
    $variables.Add('envOSVersion', [version][System.Diagnostics.FileVersionInfo]::GetVersionInfo([System.IO.Path]::Combine($variables.envSysNativeDirectory, 'ntoskrnl.exe')).ProductVersion)
    $variables.Add('envOSVersionMajor', $variables.envOSVersion.Major)
    $variables.Add('envOSVersionMinor', $variables.envOSVersion.Minor)
    $variables.Add('envOSVersionBuild', $(if ($variables.envOSVersion.Build -ge 0) { $variables.envOSVersion.Build }))
    $variables.Add('envOSVersionRevision', $(if ($variables.envOSVersion.Revision -ge 0) { $variables.envOSVersion.Revision }))

    # Get the operating system type.
    $variables.Add('envOSProductType', $variables.envOS.ProductType)
    $variables.Add('IsServerOS', $variables.envOSProductType -eq 3)
    $variables.Add('IsDomainControllerOS', $variables.envOSProductType -eq 2)
    $variables.Add('IsWorkstationOS', $variables.envOSProductType -eq 1)
    $variables.Add('IsMultiSessionOS', (& $Script:CommandTable.'Test-ADTIsMultiSessionOS'))
    $variables.Add('envOSProductTypeName', $(switch ($variables.envOSProductType)
            {
                3 { 'Server'; break }
                2 { 'Domain Controller'; break }
                1 { 'Workstation'; break }
                default { 'Unknown'; break }
            }))

    ## Variables: Office C2R version, bitness and channel
    $variables.Add('envOfficeVars', (& $Script:CommandTable.'Get-ItemProperty' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Ignore))
    $variables.Add('envOfficeVersion', ($variables.envOfficeVars | & $Script:CommandTable.'Select-Object' -ExpandProperty VersionToReport -ErrorAction Ignore))
    $variables.Add('envOfficeBitness', ($variables.envOfficeVars | & $Script:CommandTable.'Select-Object' -ExpandProperty Platform -ErrorAction Ignore))

    # Channel needs special handling for group policy values.
    $officeChannelProperty = if ($variables.envOfficeVars | & $Script:CommandTable.'Select-Object' -ExpandProperty UpdateChannel -ErrorAction Ignore)
    {
        $variables.envOfficeVars.UpdateChannel
    }
    elseif ($variables.envOfficeVars | & $Script:CommandTable.'Select-Object' -ExpandProperty CDNBaseURL -ErrorAction Ignore)
    {
        $variables.envOfficeVars.CDNBaseURL
    }
    $variables.Add('envOfficeChannel', $(switch ($officeChannelProperty -replace '^.+/')
            {
                "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" { "monthly"; break }
                "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" { "semi-annual"; break }
                "64256afe-f5d9-4f86-8936-8840a6a4f5be" { "monthly targeted"; break }
                "b8f9b850-328d-4355-9145-c59439a0c4cf" { "semi-annual targeted"; break }
                "55336b82-a18d-4dd6-b5f6-9e5095c314a6" { "monthly enterprise"; break }
            }))

    ## Variables: Hardware
    $w32b = & $Script:CommandTable.'Get-CimInstance' -ClassName Win32_BIOS -Verbose:$false
    $variables.Add('envSystemRAM', [System.Math]::Round($w32cs.TotalPhysicalMemory / 1GB))
    $variables.Add('envHardwareType', $(if (($w32b.Version -match 'VRTUAL') -or (($w32cs.Manufacturer -like '*Microsoft*') -and ($w32cs.Model -notlike '*Surface*')))
            {
                'Virtual:Hyper-V'
            }
            elseif ($w32b.Version -match 'A M I')
            {
                'Virtual:Virtual PC'
            }
            elseif ($w32b.Version -like '*Xen*')
            {
                'Virtual:Xen'
            }
            elseif (($w32b.SerialNumber -like '*VMware*') -or ($w32cs.Manufacturer -like '*VMWare*'))
            {
                'Virtual:VMware'
            }
            elseif (($w32b.SerialNumber -like '*Parallels*') -or ($w32cs.Manufacturer -like '*Parallels*'))
            {
                'Virtual:Parallels'
            }
            elseif ($w32cs.Model -like '*Virtual*')
            {
                'Virtual'
            }
            else
            {
                'Physical'
            }))

    ## Variables: PowerShell And CLR (.NET) Versions
    $variables.Add('envPSVersionTable', $PSVersionTable)
    $variables.Add('envPSProcessPath', (& $Script:CommandTable.'Get-ADTPowerShellProcessPath'))

    # PowerShell Version
    $variables.Add('envPSVersion', [System.Version]$variables.envPSVersionTable.PSVersion)
    $variables.Add('envPSVersionSemantic', $(if ($variables.envPSVersionTable.PSVersion.GetType().FullName.Equals('System.Management.Automation.SemanticVersion')) { $variables.envPSVersionTable.PSVersion }))
    $variables.Add('envPSVersionMajor', $variables.envPSVersion.Major)
    $variables.Add('envPSVersionMinor', $variables.envPSVersion.Minor)
    $variables.Add('envPSVersionBuild', $(if ($variables.envPSVersion.Build -ge 0) { $variables.envPSVersion.Build }))
    $variables.Add('envPSVersionRevision', $(if ($variables.envPSVersion.Revision -ge 0) { $variables.envPSVersion.Revision }))
    $variables.Add('envPSVersionPreReleaseLabel', $(if ($variables.envPSVersionSemantic -and $variables.envPSVersionSemantic.PreReleaseLabel) { $variables.envPSVersionSemantic.PreReleaseLabel }))
    $variables.Add('envPSVersionBuildLabel', $(if ($variables.envPSVersionSemantic -and $variables.envPSVersionSemantic.BuildLabel) { $variables.envPSVersionSemantic.BuildLabel }))

    # CLR (.NET) Version used by Windows PowerShell
    if ($variables.envPSVersionTable.ContainsKey('CLRVersion'))
    {
        $variables.Add('envCLRVersion', $variables.envPSVersionTable.CLRVersion)
        $variables.Add('envCLRVersionMajor', $variables.envCLRVersion.Major)
        $variables.Add('envCLRVersionMinor', $variables.envCLRVersion.Minor)
        $variables.Add('envCLRVersionBuild', $(if ($variables.envCLRVersion.Build -ge 0) { $variables.envCLRVersion.Build }))
        $variables.Add('envCLRVersionRevision', $(if ($variables.envCLRVersion.Revision -ge 0) { $variables.envCLRVersion.Revision }))
    }
    else
    {
        $variables.Add('envCLRVersion', $null)
        $variables.Add('envCLRVersionMajor', $null)
        $variables.Add('envCLRVersionMinor', $null)
        $variables.Add('envCLRVersionBuild', $null)
        $variables.Add('envCLRVersionRevision', $null)
    }

    ## Variables: Permissions/Accounts
    $variables.Add('CurrentProcessToken', [System.Security.Principal.WindowsIdentity]::GetCurrent())
    $variables.Add('CurrentProcessSID', [System.Security.Principal.SecurityIdentifier]$variables.CurrentProcessToken.User)
    $variables.Add('ProcessNTAccount', $variables.CurrentProcessToken.Name)
    $variables.Add('ProcessNTAccountSID', $variables.CurrentProcessSID.Value)
    $variables.Add('IsAdmin', (& $Script:CommandTable.'Test-ADTCallerIsAdmin'))
    $variables.Add('IsLocalSystemAccount', $variables.CurrentProcessSID.IsWellKnown([System.Security.Principal.WellKnownSidType]::LocalSystemSid))
    $variables.Add('IsLocalServiceAccount', $variables.CurrentProcessSID.IsWellKnown([System.Security.Principal.WellKnownSidType]::LocalServiceSid))
    $variables.Add('IsNetworkServiceAccount', $variables.CurrentProcessSID.IsWellKnown([System.Security.Principal.WellKnownSidType]::NetworkServiceSid))
    $variables.Add('IsServiceAccount', ($variables.CurrentProcessToken.Groups -contains ([System.Security.Principal.SecurityIdentifier]'S-1-5-6')))
    $variables.Add('IsProcessUserInteractive', [System.Environment]::UserInteractive)
    $variables.Add('LocalSystemNTAccount', (& $Script:CommandTable.'ConvertTo-ADTNTAccountOrSID' -WellKnownSIDName LocalSystemSid -WellKnownToNTAccount -LocalHost 4>$null).Value)
    $variables.Add('LocalUsersGroup', (& $Script:CommandTable.'ConvertTo-ADTNTAccountOrSID' -WellKnownSIDName BuiltinUsersSid -WellKnownToNTAccount -LocalHost 4>$null).Value)
    $variables.Add('LocalAdministratorsGroup', (& $Script:CommandTable.'ConvertTo-ADTNTAccountOrSID' -WellKnownSIDName BuiltinAdministratorsSid -WellKnownToNTAccount -LocalHost 4>$null).Value)
    $variables.Add('SessionZero', $variables.IsLocalSystemAccount -or $variables.IsLocalServiceAccount -or $variables.IsNetworkServiceAccount -or $variables.IsServiceAccount)

    ## Variables: Logged on user information
    $variables.Add('LoggedOnUserSessions', (& $Script:CommandTable.'Get-ADTLoggedOnUser'))
    $variables.Add('usersLoggedOn', ($variables.LoggedOnUserSessions | & { process { if ($_) { $_.NTAccount } } }))
    $variables.Add('CurrentLoggedOnUserSession', ($variables.LoggedOnUserSessions | & { process { if ($_ -and $_.IsCurrentSession) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1))
    $variables.Add('CurrentConsoleUserSession', ($variables.LoggedOnUserSessions | & { process { if ($_ -and $_.IsConsoleSession) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1))
    $variables.Add('RunAsActiveUser', $(if ($null -ne $variables.LoggedOnUserSessions) { & $Script:CommandTable.'Get-ADTRunAsActiveUser' -UserSessionInfo $variables.LoggedOnUserSessions }))

    ## Variables: User profile information.
    $variables.Add('dirUserProfile', [System.IO.Directory]::GetParent($variables.envPublic))
    $variables.Add('userProfileName', $(if ($variables.RunAsActiveUser) { $variables.RunAsActiveUser.UserName }))
    $variables.Add('runasUserProfile', $(if ($variables.userProfileName) { & $Script:CommandTable.'Join-Path' -Path $variables.dirUserProfile -ChildPath $variables.userProfileName -Resolve -ErrorAction Ignore }))

    ## Variables: Invalid FileName Characters
    $variables.Add('invalidFileNameChars', [System.IO.Path]::GetInvalidFileNameChars())
    $variables.Add('invalidFileNameCharsRegExPattern', [System.Text.RegularExpressions.Regex]::Escape([System.String]::Join($null, $variables.invalidFileNameChars)))

    ## Variables: RegEx Patterns
    $variables.Add('MSIProductCodeRegExPattern', '^(\{{0,1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{0,1})$')
    $variables.Add('InvalidScheduledTaskNameCharsRegExPattern', "[$([System.Text.RegularExpressions.Regex]::Escape('\/:*?"<>|'))]")

    # Add in WScript shell variables.
    $variables.Add('Shell', [System.Activator]::CreateInstance([System.Type]::GetTypeFromProgID('WScript.Shell')))
    $variables.Add('ShellApp', [System.Activator]::CreateInstance([System.Type]::GetTypeFromProgID('Shell.Application')))

    # Return variables for use within the module.
    return $variables.AsReadOnly()
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTPreferenceVariables
#
#-----------------------------------------------------------------------------

function Set-ADTPreferenceVariables
{
    <#
    .SYNOPSIS
        Sets preference variables within the called scope based on CommonParameter values within the callstack.

    .DESCRIPTION
        Script module functions do not automatically inherit their caller's variables, therefore we walk the callstack to get the closest bound CommonParameter value and use it within the called scope.

        This function is a helper function for any script module Advanced Function; by passing in the values of $ExecutionContext.SessionState, Set-ADTPreferenceVariables will set the caller's preference variables locally.

    .PARAMETER SessionState
        The $ExecutionContext.SessionState object from a script module Advanced Function. This is how the Set-ADTPreferenceVariables function sets variables in its callers' scope, even if that caller is in a different script module.

    .PARAMETER Scope
        A scope override, mostly so this can be called via Initialize-ADTFunction.

    .EXAMPLE
        Set-ADTPreferenceVariables -SessionState $ExecutionContext.SessionState

        Imports the default PowerShell preference variables from the caller into the local scope.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .NOTES
        An active ADT session is required to use this function.

        Original code inspired by: https://gallery.technet.microsoft.com/scriptcenter/Inherit-Preference-82343b9d

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This compatibility wrapper function cannot have its name changed for backwards compatiblity purposes.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SessionState]$SessionState,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$Scope = 1
    )

    # Get the callstack so we can enumerate bound parameters of our callers.
    $stackParams = (& $Script:CommandTable.'Get-PSCallStack').InvocationInfo.BoundParameters.GetEnumerator().GetEnumerator()

    # Loop through each common parameter and get the first bound value.
    foreach ($pref in $Script:PreferenceVariableTable.GetEnumerator())
    {
        # Return early if we have nothing.
        if (!($param = $stackParams | & { process { if ($_.Key.Equals($pref.Key)) { return @{ Name = $pref.Value; Value = $_.Value } } } } | & $Script:CommandTable.'Select-Object' -First 1))
        {
            continue
        }

        # If we've hit a switch, default it to an ActionPreference of Continue.
        if ($param.Value -is [System.Management.Automation.SwitchParameter])
        {
            if (!$param.Value)
            {
                continue
            }
            $param.Value = [System.Management.Automation.ActionPreference]::Continue
        }

        # When we're within the same module, just go up a scope level to set the value.
        # If the caller in an external scope, we set this within their SessionState.
        if ($SessionState.Equals($ExecutionContext.SessionState))
        {
            & $Script:CommandTable.'Set-Variable' @param -Scope $Scope -Force -Confirm:$false -WhatIf:$false
        }
        else
        {
            $SessionState.PSVariable.Set($param.Value, $param.Value)
        }
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTHelpConsoleInternal
#
#-----------------------------------------------------------------------------

function Show-ADTHelpConsoleInternal
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ModuleName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Guid]$Guid,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Version]$ModuleVersion
    )

    # Ensure script runs in strict mode since this may be called in a new scope.
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    Set-StrictMode -Version 3

    # Import the module and store its passthru data so we can access it later.
    $module = Import-Module -FullyQualifiedName ([Microsoft.PowerShell.Commands.ModuleSpecification]::new($PSBoundParameters)) -PassThru

    # Build out the form's listbox.
    $helpListBox = [System.Windows.Forms.ListBox]::new()
    $helpListBox.ClientSize = [System.Drawing.Size]::new(261, 675)
    $helpListBox.Font = [System.Drawing.SystemFonts]::MessageBoxFont
    $helpListBox.Location = [System.Drawing.Point]::new(3, 0)
    $helpListBox.add_SelectedIndexChanged({ $helpTextBox.Text = [System.String]::Join("`n", ((Get-Help -Name $helpListBox.SelectedItem -Full | Out-String -Stream -Width ([System.Int32]::MaxValue)) -replace '^\s+$').TrimEnd()).Trim() })
    $null = $helpListBox.Items.AddRange(($module.ExportedCommands.Keys | Sort-Object))

    # Build out the form's textbox.
    $helpTextBox = [System.Windows.Forms.RichTextBox]::new()
    $helpTextBox.ClientSize = [System.Drawing.Size]::new(1250, 675)
    $helpTextBox.Font = [System.Drawing.Font]::new('Consolas', 9)
    $helpTextBox.Location = [System.Drawing.Point]::new(271, 0)
    $helpTextBox.ReadOnly = $true
    $helpTextBox.WordWrap = $false

    # Build out the form. The suspend/resume is crucial for HiDPI support!
    $helpForm = [System.Windows.Forms.Form]::new()
    $helpForm.SuspendLayout()
    $helpForm.Text = "$($module.Name) Help Console"
    $helpForm.Font = [System.Drawing.SystemFonts]::MessageBoxFont
    $helpForm.AutoScaleDimensions = [System.Drawing.SizeF]::new(7, 15)
    $helpForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $helpForm.AutoSize = $true
    $helpForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $helpForm.MaximizeBox = $false
    $helpForm.Controls.Add($helpListBox)
    $helpForm.Controls.Add($helpTextBox)
    $helpForm.ResumeLayout()

    # Show the form. Using Application.Run automatically manages disposal for us.
    [System.Windows.Forms.Application]::Run($helpForm)
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationProgressClassic
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationProgressClassic
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'NoRelocation', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WindowTitle,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$StatusMessage,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Default', 'TopLeft', 'Top', 'TopRight', 'TopCenter', 'BottomLeft', 'Bottom', 'BottomRight')]
        [System.String]$WindowLocation = 'Default',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.TextAlignment]$MessageAlignment = [System.Windows.TextAlignment]::Center,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoRelocation,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Internal worker function.
    function Update-WindowLocation
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This is an internal worker function that requires no end user confirmation.')]
        [CmdletBinding(SupportsShouldProcess = $false)]
        param
        (
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.Windows.Window]$Window,

            [Parameter(Mandatory = $false)]
            [ValidateSet('Default', 'TopLeft', 'Top', 'TopRight', 'TopCenter', 'BottomLeft', 'Bottom', 'BottomRight')]
            [System.String]$Location = 'Default'
        )

        # Calculate the position on the screen where the progress dialog should be placed.
        [System.Double]$screenCenterWidth = [System.Windows.SystemParameters]::WorkArea.Width - $Window.ActualWidth
        [System.Double]$screenCenterHeight = [System.Windows.SystemParameters]::WorkArea.Height - $Window.ActualHeight

        # Set the start position of the Window based on the screen size.
        switch ($Location)
        {
            'TopLeft'
            {
                $Window.Left = 0.
                $Window.Top = 0.
                break
            }
            'Top'
            {
                $Window.Left = $screenCenterWidth * 0.5
                $Window.Top = 0.
                break
            }
            'TopRight'
            {
                $Window.Left = $screenCenterWidth
                $Window.Top = 0.
                break
            }
            'TopCenter'
            {
                $Window.Left = $screenCenterWidth * 0.5
                $Window.Top = $screenCenterHeight * (1. / 6.)
                break
            }
            'BottomLeft'
            {
                $Window.Left = 0.
                $Window.Top = $screenCenterHeight
                break
            }
            'Bottom'
            {
                $Window.Left = $screenCenterWidth * 0.5
                $Window.Top = $screenCenterHeight
                break
            }
            'BottomRight'
            {
                # The -100 offset is needed to not overlap system tray toast notifications.
                $Window.Left = $screenCenterWidth
                $Window.Top = $screenCenterHeight - 100
                break
            }
            default
            {
                # Center the progress window by calculating the center of the workable screen based on the width of the screen minus half the width of the progress bar
                $Window.Left = $screenCenterWidth * 0.5
                $Window.Top = $screenCenterHeight * 0.5
                break
            }
        }
    }

    # Check if the progress thread is running before invoking methods on it.
    if (!$Script:Dialogs.Classic.ProgressWindow.Running)
    {
        # Load up the XML file.
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
        $xaml = [System.Xml.XmlDocument]::new()
        $xaml.Load($Script:Dialogs.Classic.ProgressWindow.XamlCode)
        $xaml.Window.Title = $xaml.Window.ToolTip = $WindowTitle
        $xaml.Window.TopMost = (!$NotTopMost).ToString()
        $xaml.Window.Grid.TextBlock.Text = $StatusMessage
        $xaml.Window.Grid.TextBlock.TextAlignment = $MessageAlignment.ToString()

        # Set up the PowerShell instance and commence invocation.
        $Script:Dialogs.Classic.ProgressWindow.PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($Script:CommandTable.'Show-ADTInstallationProgressClassicInternal'.ScriptBlock).AddArgument($Xaml).AddArgument($adtConfig.Assets.Logo).AddArgument($adtConfig.Assets.Banner).AddArgument($WindowLocation).AddArgument(${Function:Update-WindowLocation}.Ast.Body.GetScriptBlock()).AddArgument($Script:CommandTable.'Disable-ADTWindowCloseButton'.ScriptBlock.Ast.Body.GetScriptBlock())
        $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.Open()
        $Script:Dialogs.Classic.ProgressWindow.PowerShell.Runspace.SessionStateProxy.SetVariable('SyncHash', $Script:Dialogs.Classic.ProgressWindow.SyncHash)
        $Script:Dialogs.Classic.ProgressWindow.Invocation = $Script:Dialogs.Classic.ProgressWindow.PowerShell.BeginInvoke()

        # Allow the thread to be spun up safely before invoking actions against it.
        while (!($Script:Dialogs.Classic.ProgressWindow.SyncHash.ContainsKey('Window') -and $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window.IsInitialized -and $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window.Dispatcher.Thread.ThreadState.Equals([System.Threading.ThreadState]::Running)))
        {
            if ($Script:Dialogs.Classic.ProgressWindow.SyncHash.ContainsKey('Error'))
            {
                $PSCmdlet.ThrowTerminatingError($Script:Dialogs.Classic.ProgressWindow.SyncHash.Error)
            }
            elseif ($Script:Dialogs.Classic.ProgressWindow.Invocation.IsCompleted)
            {
                $naerParams = @{
                    Exception = [System.InvalidOperationException]::new("The separate thread completed without presenting the progress dialog.")
                    Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                    ErrorId = 'InstallationProgressDialogFailure'
                    TargetObject = $(if ($Script:Dialogs.Classic.ProgressWindow.SyncHash.ContainsKey('Window')) { $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window })
                    RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                }
                $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
            }
        }

        # If we're here, the window came up.
        $Script:Dialogs.Classic.ProgressWindow.Running = $true
    }
    else
    {
        # Invoke update events against an established window.
        $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window.Dispatcher.Invoke(
            {
                $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window.Title = $WindowTitle
                $Script:Dialogs.Classic.ProgressWindow.SyncHash.Message.Text = $StatusMessage
                $Script:Dialogs.Classic.ProgressWindow.SyncHash.Message.TextAlignment = $MessageAlignment
                if (!$NoRelocation)
                {
                    Update-WindowLocation -Window $Script:Dialogs.Classic.ProgressWindow.SyncHash.Window -Location $WindowLocation
                }
            },
            [System.Windows.Threading.DispatcherPriority]::Send
        )
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationProgressClassicInternal
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationProgressClassicInternal
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'DisableWindowCloseButton', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UpdateWindowLocation', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'WindowLocation', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Xml.XmlDocument]$Xaml,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo]$Icon,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo]$Banner,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Default', 'TopLeft', 'Top', 'TopRight', 'TopCenter', 'BottomLeft', 'Bottom', 'BottomRight')]
        [System.String]$WindowLocation,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]$UpdateWindowLocation,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]$DisableWindowCloseButton
    )

    # Set required variables to ensure script functionality.
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    Set-StrictMode -Version 3

    # Create XAML window and bring it up.
    try
    {
        $SyncHash.Add('Window', [System.Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($Xaml)))
        $SyncHash.Add('Message', $SyncHash.Window.FindName('ProgressText'))
        $SyncHash.Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([System.IO.MemoryStream]::new([System.IO.File]::ReadAllBytes($Icon)), [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $SyncHash.Window.FindName('ProgressBanner').Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([System.IO.MemoryStream]::new([System.IO.File]::ReadAllBytes($Banner)), [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $SyncHash.Window.add_MouseLeftButtonDown({ $this.DragMove() })
        $SyncHash.Window.add_Loaded({
                # Relocate the window and disable the X button.
                & $UpdateWindowLocation -Window $this -Location $WindowLocation
                & $DisableWindowCloseButton -WindowHandle ([System.Windows.Interop.WindowInteropHelper]::new($this).Handle)
            })
        $null = $SyncHash.Window.ShowDialog()
    }
    catch
    {
        $SyncHash.Add('Error', $_)
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationProgressFluent
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationProgressFluent
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WindowTitle,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WindowSubtitle,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$StatusMessage,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$StatusMessageDetail,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Perform initial setup.
    $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

    # Advise that repositioning the progress window is unsupported for fluent.
    if ($UnboundArguments -eq '-WindowLocation:')
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "The parameter [-WindowLocation] is not supported with fluent dialogs and has no effect." -Severity 2
    }

    # Check if the progress thread is running before invoking methods on it.
    if (!$Script:Dialogs.Fluent.ProgressWindow.Running)
    {
        # Instantiate a new progress window object and start it up.
        [PSADT.UserInterface.UnifiedADTApplication]::ShowProgressDialog(
            $WindowTitle,
            $WindowSubtitle,
            !$NotTopMost,
            $adtConfig.Assets.Logo,
            $StatusMessage,
            $StatusMessageDetail
        )

        # Allow the thread to be spun up safely before invoking actions against it.
        do
        {
            $Script:Dialogs.Fluent.ProgressWindow.Running = [PSADT.UserInterface.UnifiedADTApplication]::CurrentDialogVisible()
        }
        until ($Script:Dialogs.Fluent.ProgressWindow.Running)
    }
    else
    {
        # Update all values.
        [PSADT.UserInterface.UnifiedADTApplication]::UpdateProgress($null, $StatusMessage, $StatusMessageDetail)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationPromptClassic
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationPromptClassic
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'formInstallationPromptStartLocation', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Left', 'Center', 'Right')]
        [System.String]$MessageAlignment = 'Center',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonRightText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonLeftText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonMiddleText,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Application', 'Asterisk', 'Error', 'Exclamation', 'Hand', 'Information', 'Question', 'Shield', 'Warning', 'WinLogo')]
        [System.String]$Icon,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PersistPrompt,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$MinimizeWindows,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$Timeout,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoExitOnTimeout,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Set up some default values.
    $controlSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, 0)
    $paddingNone = [System.Windows.Forms.Padding]::new(0, 0, 0, 0)
    $buttonSize = [System.Drawing.Size]::new(130, 24)
    $adtEnv = & $Script:CommandTable.'Get-ADTEnvironmentTable'
    $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

    # Define events for form windows.
    $installPromptTimer_Tick = {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Installation action not taken within a reasonable amount of time.'
        $buttonAbort.PerformClick()
    }
    $installPromptTimerPersist_Tick = {
        $formInstallationPrompt.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $formInstallationPrompt.TopMost = !$NotTopMost
        $formInstallationPrompt.Location = $formInstallationPromptStartLocation
        $formInstallationPrompt.BringToFront()
    }
    $formInstallationPrompt_FormClosed = {
        # Remove all event handlers from the controls.
        $installPromptTimer.remove_Tick($installPromptTimer_Tick)
        $installPromptTimer.Dispose()
        $installPromptTimer = $null
        $installPromptTimerPersist.remove_Tick($installPromptTimerPersist_Tick)
        $installPromptTimerPersist.Dispose()
        $installPromptTimerPersist = $null
        $formInstallationPrompt.remove_Load($formInstallationPrompt_Load)
        $formInstallationPrompt.remove_FormClosed($formInstallationPrompt_FormClosed)
        $formInstallationPrompt.Dispose()
        $formInstallationPrompt = $null
    }
    $formInstallationPrompt_Load = {
        # Disable the X button.
        try
        {
            & $Script:CommandTable.'Disable-ADTWindowCloseButton' -WindowHandle $formInstallationPrompt.Handle
        }
        catch
        {
            # Not a terminating error if we can't disable the button. Just disable the Control Box instead.
            & $Script:CommandTable.'Write-ADTLogEntry' 'Failed to disable the Close button. Disabling the Control Box instead.' -Severity 2
            $formInstallationPrompt.ControlBox = $false
        }

        # Correct the initial state of the form to prevent the .NET maximized form issue.
        $formInstallationPrompt.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $formInstallationPrompt.BringToFront()

        # Get the start position of the form so we can return the form to this position if PersistPrompt is enabled.
        $formInstallationPromptStartLocation = $formInstallationPrompt.Location
    }

    # Built out timer
    $installPromptTimer = [System.Windows.Forms.Timer]::new()
    $installPromptTimer.Interval = $Timeout * 1000
    $installPromptTimer.add_Tick($installPromptTimer_Tick)

    # Built out timer for Persist Prompt mode.
    $installPromptTimerPersist = [System.Windows.Forms.Timer]::new()
    $installPromptTimerPersist.Interval = $adtConfig.UI.DefaultPromptPersistInterval * 1000
    $installPromptTimerPersist.add_Tick($installPromptTimerPersist_Tick)

    # Picture Banner.
    $pictureBanner = [System.Windows.Forms.PictureBox]::new()
    $pictureBanner.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pictureBanner.MinimumSize = $pictureBanner.ClientSize = $pictureBanner.MaximumSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, $Script:Dialogs.Classic.BannerHeight)
    $pictureBanner.Location = [System.Drawing.Point]::new(0, 0)
    $pictureBanner.Name = 'PictureBanner'
    $pictureBanner.Image = $Script:Dialogs.Classic.Assets.Banner
    $pictureBanner.Margin = $paddingNone
    $pictureBanner.TabStop = $false

    # Label Text.
    $labelMessage = [System.Windows.Forms.Label]::new()
    $labelMessage.MinimumSize = $labelMessage.ClientSize = $labelMessage.MaximumSize = [System.Drawing.Size]::new(381, 0)
    $labelMessage.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 5)
    $labelMessage.Padding = [System.Windows.Forms.Padding]::new(20, 0, 20, 0)
    $labelMessage.Anchor = [System.Windows.Forms.AnchorStyles]::None
    $labelMessage.Font = $Script:Dialogs.Classic.Font
    $labelMessage.Name = 'LabelMessage'
    $labelMessage.Text = $Message
    $labelMessage.TextAlign = [System.Drawing.ContentAlignment]::"Middle$MessageAlignment"
    $labelMessage.TabStop = $false
    $labelMessage.AutoSize = $true

    # Picture Icon.
    if ($Icon)
    {
        $pictureIcon = [System.Windows.Forms.PictureBox]::new()
        $pictureIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
        $pictureIcon.MinimumSize = $pictureIcon.ClientSize = $pictureIcon.MaximumSize = [System.Drawing.Size]::new(64, 32)
        $pictureIcon.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 5)
        $pictureIcon.Padding = [System.Windows.Forms.Padding]::new(24, 0, 8, 0)
        $pictureIcon.Anchor = [System.Windows.Forms.AnchorStyles]::None
        $pictureIcon.Name = 'PictureIcon'
        $pictureIcon.Image = ([System.Drawing.SystemIcons]::$Icon).ToBitmap()
        $pictureIcon.TabStop = $false
        $pictureIcon.Height = $labelMessage.Height
    }

    # Button Abort (Hidden).
    $buttonAbort = [System.Windows.Forms.Button]::new()
    $buttonAbort.MinimumSize = $buttonAbort.ClientSize = $buttonAbort.MaximumSize = [System.Drawing.Size]::new(0, 0)
    $buttonAbort.Margin = $buttonAbort.Padding = $paddingNone
    $buttonAbort.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $buttonAbort.Name = 'ButtonAbort'
    $buttonAbort.Font = $Script:Dialogs.Classic.Font
    $buttonAbort.BackColor = [System.Drawing.Color]::Transparent
    $buttonAbort.ForeColor = [System.Drawing.Color]::Transparent
    $buttonAbort.FlatAppearance.BorderSize = 0
    $buttonAbort.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
    $buttonAbort.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
    $buttonAbort.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $buttonAbort.TabStop = $false
    $buttonAbort.Visible = $true  # Has to be set visible so we can call Click on it.
    $buttonAbort.UseVisualStyleBackColor = $true

    # Button Default (Hidden).
    $buttonDefault = [System.Windows.Forms.Button]::new()
    $buttonDefault.MinimumSize = $buttonDefault.ClientSize = $buttonDefault.MaximumSize = [System.Drawing.Size]::new(0, 0)
    $buttonDefault.Margin = $buttonDefault.Padding = $paddingNone
    $buttonDefault.Name = 'buttonDefault'
    $buttonDefault.Font = $Script:Dialogs.Classic.Font
    $buttonDefault.BackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.ForeColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatAppearance.BorderSize = 0
    $buttonDefault.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $buttonDefault.TabStop = $false
    $buttonDefault.Enabled = $false
    $buttonDefault.Visible = $true  # Has to be set visible so we can call Click on it.
    $buttonDefault.UseVisualStyleBackColor = $true

    # FlowLayoutPanel.
    $flowLayoutPanel = [System.Windows.Forms.FlowLayoutPanel]::new()
    $flowLayoutPanel.SuspendLayout()
    $flowLayoutPanel.MinimumSize = $flowLayoutPanel.ClientSize = $flowLayoutPanel.MaximumSize = $controlSize
    $flowLayoutPanel.Location = [System.Drawing.Point]::new(0, $Script:Dialogs.Classic.BannerHeight)
    $flowLayoutPanel.AutoSize = $true
    $flowLayoutPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $flowLayoutPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $flowLayoutPanel.WrapContents = $true
    $flowLayoutPanel.Margin = $flowLayoutPanel.Padding = $paddingNone

    # Make sure label text is positioned correctly before adding it.
    if ($Icon)
    {
        $labelMessage.Padding = [System.Windows.Forms.Padding]::new(0, 0, 10, 0)
        $labelMessage.Location = [System.Drawing.Point]::new(64, 0)
        $pictureIcon.Location = [System.Drawing.Point]::new(0, 0)
        $flowLayoutPanel.Controls.Add($pictureIcon)
    }
    else
    {
        $labelMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelMessage.Location = [System.Drawing.Point]::new(0, 0)
        $labelMessage.MinimumSize = $labelMessage.ClientSize = $labelMessage.MaximumSize = $controlSize
    }
    $flowLayoutPanel.Controls.Add($labelMessage)

    # Add in remaining controls and resume object.
    if ($ButtonLeftText -or $ButtonMiddleText -or $ButtonRightText)
    {
        # ButtonsPanel.
        $panelButtons = [System.Windows.Forms.Panel]::new()
        $panelButtons.SuspendLayout()
        $panelButtons.MinimumSize = $panelButtons.ClientSize = $panelButtons.MaximumSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, 39)
        $panelButtons.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
        $panelButtons.AutoSize = $true
        if ($Icon)
        {
            $panelButtons.Location = [System.Drawing.Point]::new(64, 0)
        }
        else
        {
            $panelButtons.Padding = $paddingNone
        }

        # Build out and add the buttons if we have any.
        if ($ButtonLeftText)
        {
            # Button Left.
            $buttonLeft = [System.Windows.Forms.Button]::new()
            $buttonLeft.MinimumSize = $buttonLeft.ClientSize = $buttonLeft.MaximumSize = $buttonSize
            $buttonLeft.Margin = $buttonLeft.Padding = $paddingNone
            $buttonLeft.Location = [System.Drawing.Point]::new(14, 4)
            $buttonLeft.DialogResult = [System.Windows.Forms.DialogResult]::No
            $buttonLeft.Font = $Script:Dialogs.Classic.Font
            $buttonLeft.Name = 'ButtonLeft'
            $buttonLeft.Text = $ButtonLeftText
            $buttonLeft.TabIndex = 0
            $buttonLeft.AutoSize = $false
            $buttonLeft.UseVisualStyleBackColor = $true
            $panelButtons.Controls.Add($buttonLeft)
        }
        if ($ButtonMiddleText)
        {
            # Button Middle.
            $buttonMiddle = [System.Windows.Forms.Button]::new()
            $buttonMiddle.MinimumSize = $buttonMiddle.ClientSize = $buttonMiddle.MaximumSize = $buttonSize
            $buttonMiddle.Margin = $buttonMiddle.Padding = $paddingNone
            $buttonMiddle.Location = [System.Drawing.Point]::new(160, 4)
            $buttonMiddle.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
            $buttonMiddle.Font = $Script:Dialogs.Classic.Font
            $buttonMiddle.Name = 'ButtonMiddle'
            $buttonMiddle.Text = $ButtonMiddleText
            $buttonMiddle.TabIndex = 1
            $buttonMiddle.AutoSize = $false
            $buttonMiddle.UseVisualStyleBackColor = $true
            $panelButtons.Controls.Add($buttonMiddle)
        }
        if ($ButtonRightText)
        {
            # Button Right.
            $buttonRight = [System.Windows.Forms.Button]::new()
            $buttonRight.MinimumSize = $buttonRight.ClientSize = $buttonRight.MaximumSize = $buttonSize
            $buttonRight.Margin = $buttonRight.Padding = $paddingNone
            $buttonRight.Location = [System.Drawing.Point]::new(306, 4)
            $buttonRight.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            $buttonRight.Font = $Script:Dialogs.Classic.Font
            $buttonRight.Name = 'ButtonRight'
            $buttonRight.Text = $ButtonRightText
            $buttonRight.TabIndex = 2
            $buttonRight.AutoSize = $false
            $buttonRight.UseVisualStyleBackColor = $true
            $panelButtons.Controls.Add($buttonRight)
        }

        # Add the button panel in if we have buttons.
        if ($panelButtons.Controls.Count)
        {
            $panelButtons.ResumeLayout()
            $flowLayoutPanel.Controls.Add($panelButtons)
        }
    }
    $flowLayoutPanel.ResumeLayout()

    # Form Installation Prompt.
    $formInstallationPromptStartLocation = $null
    $formInstallationPrompt = [System.Windows.Forms.Form]::new()
    $formInstallationPrompt.SuspendLayout()
    $formInstallationPrompt.ClientSize = $controlSize
    $formInstallationPrompt.Margin = $formInstallationPrompt.Padding = $paddingNone
    $formInstallationPrompt.Font = $Script:Dialogs.Classic.Font
    $formInstallationPrompt.Name = 'InstallPromptForm'
    $formInstallationPrompt.Text = $Title
    $formInstallationPrompt.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $formInstallationPrompt.AutoScaleDimensions = [System.Drawing.SizeF]::new(7, 15)
    $formInstallationPrompt.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $formInstallationPrompt.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $formInstallationPrompt.MaximizeBox = $false
    $formInstallationPrompt.MinimizeBox = $false
    $formInstallationPrompt.TopMost = !$NotTopMost
    $formInstallationPrompt.TopLevel = $true
    $formInstallationPrompt.AutoSize = $true
    $formInstallationPrompt.Icon = $Script:Dialogs.Classic.Assets.Icon
    $formInstallationPrompt.Controls.Add($pictureBanner)
    $formInstallationPrompt.Controls.Add($buttonAbort)
    $formInstallationPrompt.Controls.Add($buttonDefault)
    $formInstallationPrompt.Controls.Add($flowLayoutPanel)
    $formInstallationPrompt.add_Load($formInstallationPrompt_Load)
    $formInstallationPrompt.add_FormClosed($formInstallationPrompt_FormClosed)
    $formInstallationPrompt.AcceptButton = $buttonDefault
    $formInstallationPrompt.ActiveControl = $buttonDefault
    $formInstallationPrompt.ResumeLayout()

    # Start the timer.
    $installPromptTimer.Start()
    if ($PersistPrompt) { $installPromptTimerPersist.Start() }

    # Show the prompt synchronously. If user cancels, then keep showing it until user responds using one of the buttons.
    do
    {
        # Minimize all other windows
        if ($MinimizeWindows)
        {
            $null = $adtEnv.ShellApp.MinimizeAll()
        }

        # Show the Form
        $formResult = $formInstallationPrompt.ShowDialog()
    }
    until ($formResult -match '^(Yes|No|Ignore|Abort)$')

    # Return the button text to the caller.
    switch ($formResult)
    {
        Yes
        {
            return $ButtonRightText
        }
        No
        {
            return $ButtonLeftText
        }
        Ignore
        {
            return $ButtonMiddleText
        }
        Abort
        {
            # Restore minimized windows.
            if ($MinimizeWindows)
            {
                $null = $adtEnv.ShellApp.UndoMinimizeAll()
            }
            if (!$NoExitOnTimeout)
            {
                if (& $Script:CommandTable.'Test-ADTSessionActive')
                {
                    & $Script:CommandTable.'Close-ADTSession' -ExitCode $adtConfig.UI.DefaultExitCode
                }
            }
            else
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'UI timed out but -NoExitOnTimeout specified. Continue...'
            }
            break
        }
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationPromptFluent
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationPromptFluent
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Subtitle,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonRightText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonLeftText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonMiddleText,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$Timeout,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Send this straight out to the C# backend.
    return [PSADT.UserInterface.UnifiedADTApplication]::ShowCustomDialog(
        [System.TimeSpan]::FromSeconds($Timeout),
        $Title,
        $Subtitle,
        !$NotTopMost,
        (& $Script:CommandTable.'Get-ADTConfig').Assets.Logo,
        $Message,
        $ButtonLeftText,
        $ButtonMiddleText,
        $ButtonRightText
    )
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationRestartPromptClassic
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationRestartPromptClassic
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'formRestartPromptStartLocation', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'CountdownNoHideSeconds', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$CountdownSeconds,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$CountdownNoHideSeconds,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Initialize variables.
    $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
    $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'

    # Define starting counters.
    $startTime = [System.DateTime]::Now
    $countdownTime = $startTime

    # Set up some default values.
    $controlSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, 0)
    $paddingNone = [System.Windows.Forms.Padding]::new(0, 0, 0, 0)
    $buttonSize = [System.Drawing.Size]::new(195, 24)

    # Define events for form windows.
    $formRestart_Load = {
        # Disable the X button.
        try
        {
            & $Script:CommandTable.'Disable-ADTWindowCloseButton' -WindowHandle $formRestart.Handle
        }
        catch
        {
            # Not a terminating error if we can't disable the button. Just disable the Control Box instead
            & $Script:CommandTable.'Write-ADTLogEntry' 'Failed to disable the Close button. Disabling the Control Box instead.' -Severity 2
            $formRestart.ControlBox = $false
        }

        # Initialize the countdown timer.
        $currentTime = [System.DateTime]::Now
        $countdownTime = $startTime.AddSeconds($countdownSeconds)
        $timerCountdown.Start()

        # Set up the form.
        $remainingTime = $countdownTime.Subtract($currentTime)
        $labelCountdown.Text = [System.String]::Format('{0}:{1:d2}:{2:d2}', $remainingTime.Days * 24 + $remainingTime.Hours, $remainingTime.Minutes, $remainingTime.Seconds)
        if ($remainingTime.TotalSeconds -le $countdownNoHideSeconds)
        {
            $buttonRestartLater.Enabled = $false
        }

        # Correct the initial state of the form to prevent the .NET maximized form issue.
        $formRestart.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $formRestart.BringToFront()

        # Get the start position of the form so we can return the form to this position if PersistPrompt is enabled.
        $formRestartPromptStartLocation = $formRestart.Location
    }
    $restartTimerPersist_Tick = {
        # Show the Restart Popup.
        $formRestart.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $formRestart.TopMost = !$NotTopMost
        $formRestart.Location = $formRestartPromptStartLocation
        $formRestart.BringToFront()
    }
    $buttonRestartLater_Click = {
        # Minimize the form.
        $formRestart.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
        if ($NoCountdown)
        {
            # Reset the persistence timer.
            $restartTimerPersist.Stop()
            $restartTimerPersist.Start()
        }
    }
    $buttonRestartNow_Click = {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Forcefully restarting the computer...'
        & $Script:CommandTable.'Restart-Computer' -Force
    }
    $timerCountdown_Tick = {
        # Get the time information.
        $currentTime = & $Script:CommandTable.'Get-Date'
        $countdownTime = $startTime.AddSeconds($countdownSeconds)
        $remainingTime = $countdownTime.Subtract($currentTime)

        # If the countdown is complete, restart the machine.
        if ($countdownTime -le $currentTime)
        {
            $buttonRestartNow.PerformClick()
        }
        else
        {
            # Update the form.
            $labelCountdown.Text = [String]::Format('{0}:{1:d2}:{2:d2}', $remainingTime.Days * 24 + $remainingTime.Hours, $remainingTime.Minutes, $remainingTime.Seconds)
            if ($remainingTime.TotalSeconds -le $countdownNoHideSeconds)
            {
                $buttonRestartLater.Enabled = $false

                # If the form is hidden when we hit the "No Hide", bring it back up.
                If ($formRestart.WindowState.Equals([System.Windows.Forms.FormWindowState]::Minimized))
                {
                    $formRestart.WindowState = [System.Windows.Forms.FormWindowState]::Normal
                    $formRestart.TopMost = !$NotTopMost
                    $formRestart.Location = $formRestartPromptStartLocation
                    $formRestart.BringToFront()
                }
            }
        }
    }
    $formRestart_FormClosed = {
        $timerCountdown.remove_Tick($timerCountdown_Tick)
        $restartTimerPersist.remove_Tick($restartTimerPersist_Tick)
        $buttonRestartNow.remove_Click($buttonRestartNow_Click)
        $buttonRestartLater.remove_Click($buttonRestartLater_Click)
        $formRestart.remove_Load($formRestart_Load)
        $formRestart.remove_FormClosed($formRestart_FormClosed)
    }
    $formRestart_FormClosing = {
        if ($_.CloseReason -eq 'UserClosing')
        {
            $_.Cancel = $true
        }
    }

    # Persistence Timer.
    $timerCountdown = [System.Windows.Forms.Timer]::new()
    $restartTimerPersist = [System.Windows.Forms.Timer]::new()
    $restartTimerPersist.Interval = $adtConfig.UI.RestartPromptPersistInterval * 1000
    $restartTimerPersist.add_Tick($restartTimerPersist_Tick)
    if ($NoCountdown)
    {
        $restartTimerPersist.Start()
    }

    # Picture Banner.
    $pictureBanner = [System.Windows.Forms.PictureBox]::new()
    $pictureBanner.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pictureBanner.MinimumSize = $pictureBanner.ClientSize = $pictureBanner.MaximumSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, $Script:Dialogs.Classic.BannerHeight)
    $pictureBanner.Location = [System.Drawing.Point]::new(0, 0)
    $pictureBanner.Name = 'PictureBanner'
    $pictureBanner.Image = $Script:Dialogs.Classic.Assets.Banner
    $pictureBanner.Margin = $paddingNone
    $pictureBanner.TabStop = $false

    # Label Message.
    $labelMessage = [System.Windows.Forms.Label]::new()
    $labelMessage.MinimumSize = $labelMessage.ClientSize = $labelMessage.MaximumSize = $controlSize
    $labelMessage.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 5)
    $labelMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    $labelMessage.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $labelMessage.Font = $Script:Dialogs.Classic.Font
    $labelMessage.Name = 'LabelMessage'
    $labelMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelMessage.TabStop = $false
    $labelMessage.AutoSize = $true
    $labelMessage.Text = if ($NoCountdown)
    {
        $adtStrings.RestartPrompt.Message
    }
    else
    {
        "$($adtStrings.RestartPrompt.Message) $($adtStrings.RestartPrompt.MessageTime)`n`n$($adtStrings.RestartPrompt.MessageRestart)"
    }

    # Label Countdown.
    $labelCountdown = [System.Windows.Forms.Label]::new()
    $labelCountdown.MinimumSize = $labelCountdown.ClientSize = $labelCountdown.MaximumSize = $controlSize
    $labelCountdown.Margin = $paddingNone
    $labelCountdown.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    $labelCountdown.Font = [System.Drawing.Font]::new($Script:Dialogs.Classic.Font.Name, ($Script:Dialogs.Classic.Font.Size + 9), [System.Drawing.FontStyle]::Bold)
    $labelCountdown.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelCountdown.Text = '00:00:00'
    $labelCountdown.Name = 'LabelCountdown'
    $labelCountdown.TabStop = $false
    $labelCountdown.AutoSize = $true

    # Panel Flow Layout.
    $flowLayoutPanel = [System.Windows.Forms.FlowLayoutPanel]::new()
    $flowLayoutPanel.SuspendLayout()
    $flowLayoutPanel.MinimumSize = $flowLayoutPanel.ClientSize = $flowLayoutPanel.MaximumSize = $controlSize
    $flowLayoutPanel.Location = [System.Drawing.Point]::new(0, $Script:Dialogs.Classic.BannerHeight)
    $flowLayoutPanel.Margin = $flowLayoutPanel.Padding = $paddingNone
    $flowLayoutPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $flowLayoutPanel.AutoSize = $true
    $flowLayoutPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $flowLayoutPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $flowLayoutPanel.WrapContents = $true
    $flowLayoutPanel.Controls.Add($labelMessage)
    if (!$NoCountdown)
    {
        # Label Time remaining message.
        $labelTimeRemaining = [System.Windows.Forms.Label]::new()
        $labelTimeRemaining.MinimumSize = $labelTimeRemaining.ClientSize = $labelTimeRemaining.MaximumSize = $controlSize
        $labelTimeRemaining.Margin = $paddingNone
        $labelTimeRemaining.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelTimeRemaining.Anchor = [System.Windows.Forms.AnchorStyles]::Top
        $labelTimeRemaining.Font = [System.Drawing.Font]::new($Script:Dialogs.Classic.Font.Name, ($Script:Dialogs.Classic.Font.Size + 3), [System.Drawing.FontStyle]::Bold)
        $labelTimeRemaining.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelTimeRemaining.Text = $adtStrings.RestartPrompt.TimeRemaining
        $labelTimeRemaining.Name = 'LabelTimeRemaining'
        $labelTimeRemaining.TabStop = $false
        $labelTimeRemaining.AutoSize = $true
        $flowLayoutPanel.Controls.Add($labelTimeRemaining)
        $flowLayoutPanel.Controls.Add($labelCountdown)
    }

    # Button Panel.
    $panelButtons = [System.Windows.Forms.Panel]::new()
    $panelButtons.SuspendLayout()
    $panelButtons.MinimumSize = $panelButtons.ClientSize = $panelButtons.MaximumSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, 39)
    $panelButtons.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
    $panelButtons.Padding = $paddingNone
    $panelButtons.AutoSize = $true

    # Button Restart Now.
    $buttonRestartNow = [System.Windows.Forms.Button]::new()
    $buttonRestartNow.MinimumSize = $buttonRestartNow.ClientSize = $buttonRestartNow.MaximumSize = $buttonSize
    $buttonRestartNow.Location = [System.Drawing.Point]::new(14, 4)
    $buttonRestartNow.Margin = $buttonRestartNow.Padding = $paddingNone
    $buttonRestartNow.Name = 'ButtonRestartNow'
    $buttonRestartNow.Font = $Script:Dialogs.Classic.Font
    $buttonRestartNow.Text = $adtStrings.RestartPrompt.ButtonRestartNow
    $buttonRestartNow.TabIndex = 1
    $buttonRestartNow.AutoSize = $true
    $buttonRestartNow.UseVisualStyleBackColor = $true
    $buttonRestartNow.add_Click($buttonRestartNow_Click)
    $panelButtons.Controls.Add($buttonRestartNow)

    # Button Minimize.
    $buttonRestartLater = [System.Windows.Forms.Button]::new()
    $buttonRestartLater.MinimumSize = $buttonRestartLater.ClientSize = $buttonRestartLater.MaximumSize = $buttonSize
    $buttonRestartLater.Location = [System.Drawing.Point]::new(240, 4)
    $buttonRestartLater.Margin = $buttonRestartLater.Padding = $paddingNone
    $buttonRestartLater.Name = 'ButtonRestartLater'
    $buttonRestartLater.Font = $Script:Dialogs.Classic.Font
    $buttonRestartLater.Text = $adtStrings.RestartPrompt.ButtonRestartLater
    $buttonRestartLater.TabIndex = 0
    $buttonRestartLater.AutoSize = $true
    $buttonRestartLater.UseVisualStyleBackColor = $true
    $buttonRestartLater.add_Click($buttonRestartLater_Click)
    $panelButtons.Controls.Add($buttonRestartLater)
    $panelButtons.ResumeLayout()

    # Add the Buttons Panel to the flowPanel.
    $flowLayoutPanel.Controls.Add($panelButtons)
    $flowLayoutPanel.ResumeLayout()

    # Button Default (Hidden).
    $buttonDefault = [System.Windows.Forms.Button]::new()
    $buttonDefault.MinimumSize = $buttonDefault.ClientSize = $buttonDefault.MaximumSize = [System.Drawing.Size]::new(0, 0)
    $buttonDefault.Margin = $buttonDefault.Padding = $paddingNone
    $buttonDefault.Name = 'buttonDefault'
    $buttonDefault.Font = $Script:Dialogs.Classic.Font
    $buttonDefault.BackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.ForeColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatAppearance.BorderSize = 0
    $buttonDefault.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $buttonDefault.TabStop = $false
    $buttonDefault.Enabled = $false
    $buttonDefault.Visible = $true  # Has to be set visible so we can call Click on it.
    $buttonDefault.UseVisualStyleBackColor = $true

    # Form Restart.
    $formRestartPromptStartLocation = $null
    $formRestart = [System.Windows.Forms.Form]::new()
    $formRestart.SuspendLayout()
    $formRestart.ClientSize = $controlSize
    $formRestart.Margin = $formRestart.Padding = $paddingNone
    $formRestart.Font = $Script:Dialogs.Classic.Font
    $formRestart.Name = 'FormRestart'
    $formRestart.Text = $Title
    $formRestart.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $formRestart.AutoScaleDimensions = [System.Drawing.SizeF]::new(7, 15)
    $formRestart.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $formRestart.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $formRestart.MaximizeBox = $false
    $formRestart.MinimizeBox = $false
    $formRestart.TopMost = !$NotTopMost
    $formRestart.TopLevel = $true
    $formRestart.AutoSize = $true
    $formRestart.Icon = $Script:Dialogs.Classic.Assets.Icon
    $formRestart.Controls.Add($pictureBanner)
    $formRestart.Controls.Add($flowLayoutPanel)
    $formRestart.Controls.Add($buttonDefault)
    $formRestart.add_Load($formRestart_Load)
    $formRestart.add_FormClosed($formRestart_FormClosed)
    $formRestart.add_FormClosing($formRestart_FormClosing)
    $formRestart.AcceptButton = $buttonDefault
    $formRestart.ActiveControl = $buttonDefault
    $formRestart.ResumeLayout()

    # Timer Countdown.
    if (!$NoCountdown)
    {
        $timerCountdown.add_Tick($timerCountdown_Tick)
    }

    # Show the Form.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Displaying restart prompt with $(if ($NoCountdown) { 'no' } else { "a [$CountdownSeconds] second" }) countdown."
    return $formRestart.ShowDialog()
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationRestartPromptFluent
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationRestartPromptFluent
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Subtitle,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$CountdownSeconds,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Perform initial setup.
    $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
    $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'

    # Send this straight out to the C# backend.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Displaying restart prompt with $(if ($NoCountdown) { 'no' } else { "a [$CountdownSeconds] second" }) countdown."
    $result = [PSADT.UserInterface.UnifiedADTApplication]::ShowRestartDialog(
        $Title,
        $Subtitle,
        !$NotTopMost,
        $adtConfig.Assets.Logo,
        $adtStrings.RestartPrompt.TimeRemaining,
        $(if (!$NoCountdown) { [System.TimeSpan]::FromSeconds($CountdownSeconds) }),
        $adtStrings.RestartPrompt.Message,
        $adtStrings.RestartPrompt.MessageRestart,
        $adtStrings.RestartPrompt.ButtonRestartLater,
        $adtStrings.RestartPrompt.ButtonRestartNow
    )

    # Restart the computer if the button was pushed.
    if ($result.Equals('Restart'))
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Forcefully restarting the computer...'
        & $Script:CommandTable.'Restart-Computer' -Force
    }

    # Return the button's result to the caller.
    return $result
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTWelcomePromptClassic
#
#-----------------------------------------------------------------------------

function Show-ADTWelcomePromptClassic
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ProcessObjects', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.WelcomeState]$WelcomeState,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DeploymentType,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.ProcessObject[]]$ProcessObjects,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if ($_ -gt (& $Script:CommandTable.'Get-ADTConfig').UI.DefaultTimeout)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName CloseProcessesCountdown -ProvidedValue $_ -ExceptionMessage 'The close applications countdown time cannot be longer than the timeout specified in the config file.'))
                }
                return ($_ -ge 0)
            })]
        [System.Double]$CloseProcessesCountdown,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$DeferTimes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DeferDeadline,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ForceCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ForceCloseProcessesCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PersistPrompt,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$AllowDefer,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoMinimizeWindows,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CustomText,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Perform initial setup.
    $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
    $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'

    # Initialize variables.
    $countdownTime = $startTime = [System.DateTime]::Now
    $showCountdown = $false
    $showCloseProcesses = $false
    $showDeference = $false
    $persistWindow = $false

    # Initial form layout: Close Applications
    if ($WelcomeState.RunningProcessDescriptions)
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Prompting the user to close application(s) [$($WelcomeState.RunningProcessDescriptions -join ',')]..."
        $showCloseProcesses = $true
    }

    # Initial form layout: Allow Deferral
    if ($AllowDefer -and (($DeferTimes -ge 0) -or $DeferDeadline))
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'The user has the option to defer.'
        $showDeference = $true

        # Remove the Z from universal sortable date time format, otherwise it could be converted to a different time zone.
        if ($DeferDeadline)
        {
            $DeferDeadline = (& $Script:CommandTable.'Get-Date' -Date ($DeferDeadline -replace 'Z')).ToString()
        }
    }

    # If deferral is being shown and 'close apps countdown' or 'persist prompt' was specified, enable those features.
    if (!$showDeference)
    {
        if ($CloseProcessesCountdown -gt 0)
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Close applications countdown has [$CloseProcessesCountdown] seconds remaining."
            $showCountdown = $true
        }
    }
    elseif ($PersistPrompt)
    {
        $persistWindow = $true
    }

    # If 'force close apps countdown' was specified, enable that feature.
    if ($ForceCloseProcessesCountdown)
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Close applications countdown has [$CloseProcessesCountdown] seconds remaining."
        $showCountdown = $true
    }

    # If 'force countdown' was specified, enable that feature.
    if ($ForceCountdown)
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Countdown has [$CloseProcessesCountdown] seconds remaining."
        $showCountdown = $true
    }

    # Set up some default values.
    $controlSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, 0)
    $paddingNone = [System.Windows.Forms.Padding]::new(0, 0, 0, 0)
    $buttonSize = [System.Drawing.Size]::new(130, 24)

    # Add the timer if it doesn't already exist - this avoids the timer being reset if the continue button is clicked.
    if (!$WelcomeState.WelcomeTimer)
    {
        $WelcomeState.WelcomeTimer = [System.Windows.Forms.Timer]::new()
    }

    # Define all form events.
    $formWelcome_FormClosed = {
        $WelcomeState.WelcomeTimer.remove_Tick($welcomeTimer_Tick)
        $welcomeTimerPersist.remove_Tick($welcomeTimerPersist_Tick)
        $timerRunningProcesses.remove_Tick($timerRunningProcesses_Tick)
        $formWelcome.remove_Load($formWelcome_Load)
        $formWelcome.remove_FormClosed($formWelcome_FormClosed)
    }
    $formWelcome_Load = {
        # Disable the X button.
        try
        {
            & $Script:CommandTable.'Disable-ADTWindowCloseButton' -WindowHandle $formWelcome.Handle
        }
        catch
        {
            # Not a terminating error if we can't disable the button. Just disable the Control Box instead
            & $Script:CommandTable.'Write-ADTLogEntry' 'Failed to disable the Close button. Disabling the Control Box instead.' -Severity 2
            $formWelcome.ControlBox = $false
        }

        # Initialize the countdown timer.
        $currentTime = [System.DateTime]::Now
        $countdownTime = $startTime.AddSeconds($CloseProcessesCountdown)
        $WelcomeState.WelcomeTimer.Start()

        # Set up the form.
        $remainingTime = $countdownTime.Subtract($currentTime)
        $labelCountdown.Text = [System.String]::Format('{0}:{1:d2}:{2:d2}', $remainingTime.Days * 24 + $remainingTime.Hours, $remainingTime.Minutes, $remainingTime.Seconds)

        # Correct the initial state of the form to prevent the .NET maximized form issue.
        $formWelcome.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $formWelcome.BringToFront()

        # Get the start position of the form so we can return the form to this position if PersistPrompt is enabled.
        $WelcomeState.FormStartLocation = $formWelcome.Location
    }
    $welcomeTimer_Tick = if ($showCountdown)
    {
        {
            # Get the time information.
            [DateTime]$currentTime = [System.DateTime]::Now
            [DateTime]$countdownTime = $startTime.AddSeconds($CloseProcessesCountdown)
            [Timespan]$remainingTime = $countdownTime.Subtract($currentTime)
            $WelcomeState.CloseProcessesCountdown = $remainingTime.TotalSeconds

            # If the countdown is complete, close the application(s) or continue.
            if ($countdownTime -le $currentTime)
            {
                if ($ForceCountdown)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Countdown timer has elapsed. Force continue.'
                    $buttonContinue.PerformClick()
                }
                else
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Close application(s) countdown timer has elapsed. Force closing application(s).'
                    if ($buttonCloseProcesses.CanFocus)
                    {
                        $buttonCloseProcesses.PerformClick()
                    }
                    else
                    {
                        $buttonContinue.PerformClick()
                    }
                }
            }
            else
            {
                # Update the form.
                $labelCountdown.Text = [System.String]::Format('{0}:{1:d2}:{2:d2}', $remainingTime.Days * 24 + $remainingTime.Hours, $remainingTime.Minutes, $remainingTime.Seconds)
            }
        }
    }
    else
    {
        $WelcomeState.WelcomeTimer.Interval = $adtConfig.UI.DefaultTimeout * 1000
        {
            $buttonAbort.PerformClick()
        }
    }
    $welcomeTimerPersist_Tick = {
        $formWelcome.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $formWelcome.TopMost = !$NotTopMost
        $formWelcome.Location = $WelcomeState.FormStartLocation
        $formWelcome.BringToFront()
    }
    $timerRunningProcesses_Tick = {
        # Grab current list of running processes.
        $dynamicRunningProcesses = & $Script:CommandTable.'Get-ADTRunningProcesses' -ProcessObjects $ProcessObjects -InformationAction SilentlyContinue
        $dynamicRunningProcessDescriptions = $dynamicRunningProcesses | & $Script:CommandTable.'Select-Object' -ExpandProperty ProcessDescription | & $Script:CommandTable.'Sort-Object' -Unique
        $previousRunningProcessDescriptions = $WelcomeState.RunningProcessDescriptions

        # Check the previous list against what's currently running.
        if (& $Script:CommandTable.'Compare-Object' -ReferenceObject @($WelcomeState.RunningProcessDescriptions | & $Script:CommandTable.'Select-Object') -DifferenceObject @($dynamicRunningProcessDescriptions | & $Script:CommandTable.'Select-Object'))
        {
            # Update the runningProcessDescriptions variable for the next time this function runs.
            $listboxCloseProcesses.Items.Clear()
            if (($WelcomeState.RunningProcessDescriptions = $dynamicRunningProcessDescriptions))
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "The running processes have changed. Updating the apps to close: [$($WelcomeState.RunningProcessDescriptions -join ',')]..."
                $listboxCloseProcesses.Items.AddRange($WelcomeState.RunningProcessDescriptions)
            }
        }

        # If CloseProcesses processes were running when the prompt was shown, and they are subsequently detected to be closed while the form is showing, then close the form. The deferral and CloseProcesses conditions will be re-evaluated.
        if ($previousRunningProcessDescriptions)
        {
            if (!$dynamicRunningProcesses)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Previously detected running processes are no longer running.'
                $formWelcome.Dispose()
            }
        }
        elseif ($dynamicRunningProcesses)
        {
            # If CloseProcesses processes were not running when the prompt was shown, and they are subsequently detected to be running while the form is showing, then close the form for relaunch. The deferral and CloseProcesses conditions will be re-evaluated.
            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'New running processes detected. Updating the form to prompt to close the running applications.'
            $formWelcome.Dispose()
        }
    }

    # Welcome Timer.
    $WelcomeState.WelcomeTimer.add_Tick($welcomeTimer_Tick)

    # Persistence Timer.
    $welcomeTimerPersist = [System.Windows.Forms.Timer]::new()
    $welcomeTimerPersist.Interval = $adtConfig.UI.DefaultPromptPersistInterval * 1000
    $welcomeTimerPersist.add_Tick($welcomeTimerPersist_Tick)
    if ($persistWindow)
    {
        $welcomeTimerPersist.Start()
    }

    # Process Re-Enumeration Timer.
    $timerRunningProcesses = [System.Windows.Forms.Timer]::new()
    $timerRunningProcesses.Interval = $adtConfig.UI.DynamicProcessEvaluationInterval * 1000
    $timerRunningProcesses.add_Tick($timerRunningProcesses_Tick)
    if ($adtConfig.UI.DynamicProcessEvaluation)
    {
        $timerRunningProcesses.Start()
    }

    # Picture Banner.
    $pictureBanner = [System.Windows.Forms.PictureBox]::new()
    $pictureBanner.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pictureBanner.MinimumSize = $pictureBanner.ClientSize = $pictureBanner.MaximumSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, $Script:Dialogs.Classic.BannerHeight)
    $pictureBanner.Location = [System.Drawing.Point]::new(0, 0)
    $pictureBanner.Name = 'PictureBanner'
    $pictureBanner.Image = $Script:Dialogs.Classic.Assets.Banner
    $pictureBanner.Margin = $paddingNone
    $pictureBanner.TabStop = $false

    # Label Welcome Message.
    $labelWelcomeMessage = [System.Windows.Forms.Label]::new()
    $labelWelcomeMessage.MinimumSize = $labelWelcomeMessage.ClientSize = $labelWelcomeMessage.MaximumSize = $controlSize
    $labelWelcomeMessage.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
    $labelWelcomeMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    $labelWelcomeMessage.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $labelWelcomeMessage.Font = $Script:Dialogs.Classic.Font
    $labelWelcomeMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelWelcomeMessage.Text = $adtStrings.DeferPrompt.WelcomeMessage
    $labelWelcomeMessage.Name = 'LabelWelcomeMessage'
    $labelWelcomeMessage.TabStop = $false
    $labelWelcomeMessage.AutoSize = $true

    # Label App Name.
    $labelAppName = [System.Windows.Forms.Label]::new()
    $labelAppName.MinimumSize = $labelAppName.ClientSize = $labelAppName.MaximumSize = $controlSize
    $labelAppName.Margin = [System.Windows.Forms.Padding]::new(0, 5, 0, 5)
    $labelAppName.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    $labelAppName.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $labelAppName.Font = [System.Drawing.Font]::new($Script:Dialogs.Classic.Font.Name, ($Script:Dialogs.Classic.Font.Size + 3), [System.Drawing.FontStyle]::Bold)
    $labelAppName.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelAppName.Text = $Title.Replace('&', '&&')
    $labelAppName.Name = 'LabelAppName'
    $labelAppName.TabStop = $false
    $labelAppName.AutoSize = $true

    # Listbox Close Applications.
    $listBoxCloseProcesses = [System.Windows.Forms.ListBox]::new()
    $listBoxCloseProcesses.MinimumSize = $listBoxCloseProcesses.ClientSize = $listBoxCloseProcesses.MaximumSize = [System.Drawing.Size]::new(420, 100)
    $listBoxCloseProcesses.Margin = [System.Windows.Forms.Padding]::new(15, 0, 15, 0)
    $listBoxCloseProcesses.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    $listboxCloseProcesses.Font = $Script:Dialogs.Classic.Font
    $listBoxCloseProcesses.FormattingEnabled = $true
    $listBoxCloseProcesses.HorizontalScrollbar = $true
    $listBoxCloseProcesses.Name = 'ListBoxCloseProcesses'
    $listBoxCloseProcesses.TabIndex = 3
    if ($WelcomeState.RunningProcessDescriptions)
    {
        $null = $listboxCloseProcesses.Items.AddRange($WelcomeState.RunningProcessDescriptions)
    }

    # Label Countdown.
    $labelCountdown = [System.Windows.Forms.Label]::new()
    $labelCountdown.MinimumSize = $labelCountdown.ClientSize = $labelCountdown.MaximumSize = $controlSize
    $labelCountdown.Margin = $paddingNone
    $labelCountdown.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    $labelCountdown.Font = [System.Drawing.Font]::new($Script:Dialogs.Classic.Font.Name, ($Script:Dialogs.Classic.Font.Size + 9), [System.Drawing.FontStyle]::Bold)
    $labelCountdown.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelCountdown.Text = '00:00:00'
    $labelCountdown.Name = 'LabelCountdown'
    $labelCountdown.TabStop = $false
    $labelCountdown.AutoSize = $true

    # Panel Flow Layout.
    $flowLayoutPanel = [System.Windows.Forms.FlowLayoutPanel]::new()
    $flowLayoutPanel.SuspendLayout()
    $flowLayoutPanel.MinimumSize = $flowLayoutPanel.ClientSize = $flowLayoutPanel.MaximumSize = $controlSize
    $flowLayoutPanel.Location = [System.Drawing.Point]::new(0, $Script:Dialogs.Classic.BannerHeight)
    $flowLayoutPanel.Margin = $flowLayoutPanel.Padding = $paddingNone
    $flowLayoutPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $flowLayoutPanel.AutoSize = $true
    $flowLayoutPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $flowLayoutPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $flowLayoutPanel.WrapContents = $true
    $flowLayoutPanel.Controls.Add($labelWelcomeMessage)
    $flowLayoutPanel.Controls.Add($labelAppName)
    if ($CustomText -and $adtStrings.WelcomePrompt.CustomMessage)
    {
        # Label CustomMessage.
        $labelCustomMessage = [System.Windows.Forms.Label]::new()
        $labelCustomMessage.MinimumSize = $labelCustomMessage.ClientSize = $labelCustomMessage.MaximumSize = $controlSize
        $labelCustomMessage.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
        $labelCustomMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelCustomMessage.Anchor = [System.Windows.Forms.AnchorStyles]::Top
        $labelCustomMessage.Font = $Script:Dialogs.Classic.Font
        $labelCustomMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelCustomMessage.Text = $adtStrings.WelcomePrompt.CustomMessage
        $labelCustomMessage.Name = 'LabelCustomMessage'
        $labelCustomMessage.TabStop = $false
        $labelCustomMessage.AutoSize = $true
        $flowLayoutPanel.Controls.Add($labelCustomMessage)
    }
    if ($showCloseProcesses)
    {
        # Label CloseProcessesMessage.
        $labelCloseProcessesMessage = [System.Windows.Forms.Label]::new()
        $labelCloseProcessesMessage.MinimumSize = $labelCloseProcessesMessage.ClientSize = $labelCloseProcessesMessage.MaximumSize = $controlSize
        $labelCloseProcessesMessage.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
        $labelCloseProcessesMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelCloseProcessesMessage.Anchor = [System.Windows.Forms.AnchorStyles]::Top
        $labelCloseProcessesMessage.Font = $Script:Dialogs.Classic.Font
        $labelCloseProcessesMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelCloseProcessesMessage.Text = $adtStrings.ClosePrompt.Message
        $labelCloseProcessesMessage.Name = 'LabelCloseProcessesMessage'
        $labelCloseProcessesMessage.TabStop = $false
        $labelCloseProcessesMessage.AutoSize = $true
        $flowLayoutPanel.Controls.Add($labelCloseProcessesMessage)

        # Listbox Close Applications.
        $flowLayoutPanel.Controls.Add($listBoxCloseProcesses)
    }
    if ($showDeference)
    {
        # Label Defer Expiry Message.
        $labelDeferExpiryMessage = [System.Windows.Forms.Label]::new()
        $labelDeferExpiryMessage.MinimumSize = $labelDeferExpiryMessage.ClientSize = $labelDeferExpiryMessage.MaximumSize = $controlSize
        $labelDeferExpiryMessage.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
        $labelDeferExpiryMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelDeferExpiryMessage.Font = $Script:Dialogs.Classic.Font
        $labelDeferExpiryMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelDeferExpiryMessage.Text = $adtStrings.DeferPrompt.ExpiryMessage
        $labelDeferExpiryMessage.Name = 'LabelDeferExpiryMessage'
        $labelDeferExpiryMessage.TabStop = $false
        $labelDeferExpiryMessage.AutoSize = $true
        $flowLayoutPanel.Controls.Add($labelDeferExpiryMessage)

        # Label Defer Deadline.
        $labelDeferDeadline = [System.Windows.Forms.Label]::new()
        $labelDeferDeadline.MinimumSize = $labelDeferDeadline.ClientSize = $labelDeferDeadline.MaximumSize = $controlSize
        $labelDeferDeadline.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
        $labelDeferDeadline.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelDeferDeadline.Font = [System.Drawing.Font]::new($Script:Dialogs.Classic.Font.Name, $Script:Dialogs.Classic.Font.Size, [System.Drawing.FontStyle]::Bold)
        $labelDeferDeadline.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelDeferDeadline.Name = 'LabelDeferDeadline'
        $labelDeferDeadline.TabStop = $false
        $labelDeferDeadline.AutoSize = $true
        if ($DeferTimes -ge 0)
        {
            $labelDeferDeadline.Text = "$($adtStrings.DeferPrompt.RemainingDeferrals) $($DeferTimes + 1)"
        }
        if ($deferDeadline)
        {
            $labelDeferDeadline.Text = "$($adtStrings.DeferPrompt.Deadline) $deferDeadline"
        }
        $flowLayoutPanel.Controls.Add($labelDeferDeadline)

        # Label Defer Expiry Message.
        $labelDeferWarningMessage = [System.Windows.Forms.Label]::new()
        $labelDeferWarningMessage.MinimumSize = $labelDeferWarningMessage.ClientSize = $labelDeferWarningMessage.MaximumSize = $controlSize
        $labelDeferWarningMessage.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
        $labelDeferWarningMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelDeferWarningMessage.Font = $Script:Dialogs.Classic.Font
        $labelDeferWarningMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelDeferWarningMessage.Text = $adtStrings.DeferPrompt.WarningMessage
        $labelDeferWarningMessage.Name = 'LabelDeferWarningMessage'
        $labelDeferWarningMessage.TabStop = $false
        $labelDeferWarningMessage.AutoSize = $true
        $flowLayoutPanel.Controls.Add($labelDeferWarningMessage)
    }
    if ($showCountdown)
    {
        # Label CountdownMessage.
        $labelCountdownMessage = [System.Windows.Forms.Label]::new()
        $labelCountdownMessage.MinimumSize = $labelCountdownMessage.ClientSize = $labelCountdownMessage.MaximumSize = $controlSize
        $labelCountdownMessage.Margin = $paddingNone
        $labelCountdownMessage.Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
        $labelCountdownMessage.Anchor = [System.Windows.Forms.AnchorStyles]::Top
        $labelCountdownMessage.Font = [System.Drawing.Font]::new($Script:Dialogs.Classic.Font.Name, ($Script:Dialogs.Classic.Font.Size + 3), [System.Drawing.FontStyle]::Bold)
        $labelCountdownMessage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $labelCountdownMessage.Name = 'LabelCountdownMessage'
        $labelCountdownMessage.TabStop = $false
        $labelCountdownMessage.AutoSize = $true
        $labelCountdownMessage.Text = if ($ForceCountdown -or !$WelcomeState.RunningProcessDescriptions)
        {
            [System.String]::Format($adtStrings.WelcomePrompt.CountdownMessage, $adtStrings.DeploymentType.$DeploymentType)
        }
        else
        {
            $adtStrings.ClosePrompt.CountdownMessage
        }
        $flowLayoutPanel.Controls.Add($labelCountdownMessage)

        ## Label Countdown.
        $flowLayoutPanel.Controls.Add($labelCountdown)
    }

    # Panel Buttons.
    $panelButtons = [System.Windows.Forms.Panel]::new()
    $panelButtons.SuspendLayout()
    $panelButtons.MinimumSize = $panelButtons.ClientSize = $panelButtons.MaximumSize = [System.Drawing.Size]::new($Script:Dialogs.Classic.Width, 39)
    $panelButtons.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
    $panelButtons.Padding = $paddingNone
    $panelButtons.AutoSize = $true
    if ($showCloseProcesses)
    {
        # Button Close For Me.
        $buttonCloseProcesses = [System.Windows.Forms.Button]::new()
        $buttonCloseProcesses.MinimumSize = $buttonCloseProcesses.ClientSize = $buttonCloseProcesses.MaximumSize = $buttonSize
        $buttonCloseProcesses.Margin = $buttonCloseProcesses.Padding = $paddingNone
        $buttonCloseProcesses.Location = [System.Drawing.Point]::new(14, 4)
        $buttonCloseProcesses.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $buttonCloseProcesses.Font = $Script:Dialogs.Classic.Font
        $buttonCloseProcesses.Name = 'ButtonCloseProcesses'
        $buttonCloseProcesses.Text = $adtStrings.ClosePrompt.ButtonClose
        $buttonCloseProcesses.TabIndex = 1
        $buttonCloseProcesses.AutoSize = $true
        $buttonCloseProcesses.UseVisualStyleBackColor = $true
        $panelButtons.Controls.Add($buttonCloseProcesses)
    }
    if ($showDeference)
    {
        # Button Defer.
        $buttonDefer = [System.Windows.Forms.Button]::new()
        $buttonDefer.MinimumSize = $buttonDefer.ClientSize = $buttonDefer.MaximumSize = $buttonSize
        $buttonDefer.Margin = $buttonDefer.Padding = $paddingNone
        $buttonDefer.Location = [System.Drawing.Point]::new((14, 160)[$showCloseProcesses], 4)
        $buttonDefer.DialogResult = [System.Windows.Forms.DialogResult]::No
        $buttonDefer.Font = $Script:Dialogs.Classic.Font
        $buttonDefer.Name = 'ButtonDefer'
        $buttonDefer.Text = $adtStrings.ClosePrompt.ButtonDefer
        $buttonDefer.TabIndex = 0
        $buttonDefer.AutoSize = $true
        $buttonDefer.UseVisualStyleBackColor = $true
        $panelButtons.Controls.Add($buttonDefer)
    }

    # Button Continue.
    $buttonContinue = [System.Windows.Forms.Button]::new()
    $buttonContinue.MinimumSize = $buttonContinue.ClientSize = $buttonContinue.MaximumSize = $buttonSize
    $buttonContinue.Margin = $buttonContinue.Padding = $paddingNone
    $buttonContinue.Location = [System.Drawing.Point]::new(306, 4)
    $buttonContinue.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $buttonContinue.Font = $Script:Dialogs.Classic.Font
    $buttonContinue.Name = 'ButtonContinue'
    $buttonContinue.Text = $adtStrings.ClosePrompt.ButtonContinue
    $buttonContinue.TabIndex = 2
    $buttonContinue.AutoSize = $true
    $buttonContinue.UseVisualStyleBackColor = $true
    if ($showCloseProcesses)
    {
        # Add tooltip to Continue button.
        $toolTip = [System.Windows.Forms.ToolTip]::new()
        $toolTip.BackColor = [Drawing.Color]::LightGoldenrodYellow
        $toolTip.IsBalloon = $false
        $toolTip.InitialDelay = 100
        $toolTip.ReshowDelay = 100
        $toolTip.SetToolTip($buttonContinue, $adtStrings.ClosePrompt.ButtonContinueTooltip)
    }
    $panelButtons.Controls.Add($buttonContinue)
    $panelButtons.ResumeLayout()

    # Add the Buttons Panel to the flowPanel.
    $flowLayoutPanel.Controls.Add($panelButtons)
    $flowLayoutPanel.ResumeLayout()

    # Button Abort (Hidden).
    $buttonAbort = [System.Windows.Forms.Button]::new()
    $buttonAbort.MinimumSize = $buttonAbort.ClientSize = $buttonAbort.MaximumSize = [System.Drawing.Size]::new(0, 0)
    $buttonAbort.Margin = $buttonAbort.Padding = $paddingNone
    $buttonAbort.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $buttonAbort.Name = 'buttonAbort'
    $buttonAbort.Font = $Script:Dialogs.Classic.Font
    $buttonAbort.BackColor = [System.Drawing.Color]::Transparent
    $buttonAbort.ForeColor = [System.Drawing.Color]::Transparent
    $buttonAbort.FlatAppearance.BorderSize = 0
    $buttonAbort.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
    $buttonAbort.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
    $buttonAbort.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $buttonAbort.TabStop = $false
    $buttonAbort.Visible = $true  # Has to be set visible so we can call Click on it.
    $buttonAbort.UseVisualStyleBackColor = $true

    # Button Default (Hidden).
    $buttonDefault = [System.Windows.Forms.Button]::new()
    $buttonDefault.MinimumSize = $buttonDefault.ClientSize = $buttonDefault.MaximumSize = [System.Drawing.Size]::new(0, 0)
    $buttonDefault.Margin = $buttonDefault.Padding = $paddingNone
    $buttonDefault.Name = 'buttonDefault'
    $buttonDefault.Font = $Script:Dialogs.Classic.Font
    $buttonDefault.BackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.ForeColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatAppearance.BorderSize = 0
    $buttonDefault.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Transparent
    $buttonDefault.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $buttonDefault.TabStop = $false
    $buttonDefault.Enabled = $false
    $buttonDefault.Visible = $true  # Has to be set visible so we can call Click on it.
    $buttonDefault.UseVisualStyleBackColor = $true

    ## Form Welcome
    $formWelcome = [System.Windows.Forms.Form]::new()
    $formWelcome.SuspendLayout()
    $formWelcome.ClientSize = $controlSize
    $formWelcome.Margin = $formWelcome.Padding = $paddingNone
    $formWelcome.Font = $Script:Dialogs.Classic.Font
    $formWelcome.Name = 'WelcomeForm'
    $formWelcome.Text = $Title
    $formWelcome.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $formWelcome.AutoScaleDimensions = [System.Drawing.SizeF]::new(7, 15)
    $formWelcome.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $formWelcome.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $formWelcome.MaximizeBox = $false
    $formWelcome.MinimizeBox = $false
    $formWelcome.TopMost = !$NotTopMost
    $formWelcome.TopLevel = $true
    $formWelcome.AutoSize = $true
    $formWelcome.Icon = $Script:Dialogs.Classic.Assets.Icon
    $formWelcome.Controls.Add($pictureBanner)
    $formWelcome.Controls.Add($buttonAbort)
    $formWelcome.Controls.Add($buttonDefault)
    $formWelcome.Controls.Add($flowLayoutPanel)
    $formWelcome.add_Load($formWelcome_Load)
    $formWelcome.add_FormClosed($formWelcome_FormClosed)
    $formWelcome.AcceptButton = $buttonDefault
    $formWelcome.ActiveControl = $buttonDefault
    $formWelcome.ResumeLayout()

    # Minimize all other windows.
    if (!$NoMinimizeWindows)
    {
        $null = (& $Script:CommandTable.'Get-ADTEnvironmentTable').ShellApp.MinimizeAll()
    }

    # Run the form and store the result.
    $result = switch ($formWelcome.ShowDialog())
    {
        OK { 'Continue'; break }
        No { 'Defer'; break }
        Yes { 'Close'; break }
        Abort { 'Timeout'; break }
    }
    $formWelcome.Dispose()

    # Shut down the timer if its running.
    if ($adtConfig.UI.DynamicProcessEvaluation)
    {
        $timerRunningProcesses.Stop()
    }

    # Return the result to the caller.
    return $result
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTWelcomePromptFluent
#
#-----------------------------------------------------------------------------

function Show-ADTWelcomePromptFluent
{
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UnboundArguments', Justification = "This parameter is just to trap any superfluous input at the end of the function's call.")]
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.WelcomeState]$WelcomeState,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Title,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Subtitle,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$DeferTimes,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoMinimizeWindows,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    # Perform initial setup.
    $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
    $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'

    # Convert the incoming ProcessObject objects into AppProcessInfo objects.
    $appsToClose = if ($WelcomeState.RunningProcesses)
    {
        $WelcomeState.RunningProcesses | & {
            process
            {
                $_.Refresh(); if (!$_.HasExited)
                {
                    # Get icon so we can convert it into a media image for the UI.
                    $icon = try
                    {
                        [PSADT.UserInterface.Utilities.ProcessExtensions]::GetIcon($_, $true)
                    }
                    catch
                    {
                        $null = $null
                    }

                    # Instantiate and return a new AppProcessInfo object.
                    return [PSADT.UserInterface.Services.AppProcessInfo]::new(
                        $_.ProcessName,
                        $_.ProcessDescription,
                        $_.Product,
                        $_.Company,
                        $(if ($icon) { [PSADT.UserInterface.Utilities.BitmapExtensions]::ConvertToImageSource($icon.ToBitmap()) }),
                        $_.StartTime
                    )
                }
            }
        }
    }

    # Minimize all other windows.
    if (!$NoMinimizeWindows)
    {
        $null = (& $Script:CommandTable.'Get-ADTEnvironmentTable').ShellApp.MinimizeAll()
    }

    # Send this out to the C# code.
    $result = [PSADT.UserInterface.UnifiedADTApplication]::ShowWelcomeDialog(
        [System.TimeSpan]::FromSeconds($adtConfig.UI.DefaultTimeout),
        $Title,
        $Subtitle,
        !$NotTopMost,
        $(if ($PSBoundParameters.ContainsKey('DeferTimes')) { $DeferTimes + 1 }),
        $appsToClose,
        $adtConfig.Assets.Logo,
        $adtStrings.WelcomePrompt.Fluent.DialogMessage,
        $adtStrings.WelcomePrompt.Fluent.DialogMessageNoProcesses,
        $adtStrings.WelcomePrompt.Fluent.ButtonDeferRemaining,
        $adtStrings.WelcomePrompt.Fluent.ButtonLeftText,
        $adtStrings.WelcomePrompt.Fluent.ButtonRightText,
        $adtStrings.WelcomePrompt.Fluent.ButtonRightTextNoProcesses,
        $(if ($adtConfig.UI.DynamicProcessEvaluation) { [PSADT.UserInterface.Services.ProcessEvaluationService]::new() })
    )

    # Return a translated value that's compatible with the toolkit.
    switch ($result)
    {
        Continue
        {
            return 'Close'
            break
        }
        Defer
        {
            return 'Defer'
            break
        }
        Cancel
        {
            return 'Timeout'
            break
        }
        default
        {
            $naerParams = @{
                Exception = [System.InvalidOperationException]::new("The returned dialog result of [$_] is invalid and cannot be processed.")
                Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                ErrorId = "WelcomeDialogInvalidResult"
                TargetObject = $_
            }
            $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
        }
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTInstallationProgressRunning
#
#-----------------------------------------------------------------------------

function Test-ADTInstallationProgressRunning
{
    # Return the value of the global state's bool.
    return $Script:Dialogs.((& $Script:CommandTable.'Get-ADTConfig').UI.DialogStyle).ProgressWindow.Running
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTIsMultiSessionOS
#
#-----------------------------------------------------------------------------

function Test-ADTIsMultiSessionOS
{
    # The registry is significantly cheaper to query than a CIM instance.
    # https://www.jasonsamuel.com/2020/03/02/how-to-use-microsoft-wvd-windows-10-multi-session-fslogix-msix-app-attach-to-build-an-azure-powered-virtual-desktop-experience/
    return ([Microsoft.Win32.Registry]::GetValue('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'ProductName', $null) -match '^Microsoft Windows \d+ Enterprise (for Virtual Desktops|Multi-Session)$')
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTModuleIsReleaseBuild
#
#-----------------------------------------------------------------------------

function Test-ADTModuleIsReleaseBuild
{
    return $Script:Module.Compiled -and $Script:Module.Signed
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTNonNativeCaller
#
#-----------------------------------------------------------------------------

function Test-ADTNonNativeCaller
{
    return (& $Script:CommandTable.'Get-PSCallStack').Command.Contains('AppDeployToolkitMain.ps1')
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTReleaseBuildFileValidity
#
#-----------------------------------------------------------------------------

function Test-ADTReleaseBuildFileValidity
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName LiteralPath -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName LiteralPath -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String]$LiteralPath
    )

    # If we're running a release module, ensure the ps*1 files haven't been tampered with.
    if ((& $Script:CommandTable.'Test-ADTModuleIsReleaseBuild') -and ($badFiles = & $Script:CommandTable.'Get-ChildItem' @PSBoundParameters -Filter *.ps*1 -Recurse | & $Script:CommandTable.'Get-AuthenticodeSignature' | & { process { if (!$_.Status.Equals([System.Management.Automation.SignatureStatus]::Valid)) { return $_ } } }))
    {
        return $badFiles
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Unblock-ADTAppExecutionInternal
#
#-----------------------------------------------------------------------------

function Unblock-ADTAppExecutionInternal
{
    <#

    .SYNOPSIS
    Core logic used within Unblock-ADTAppExecution.

    .DESCRIPTION
    This function contains core logic used within Unblock-ADTAppExecution, separated out to facilitate calling via PowerShell without dependency on the toolkit.

    .NOTES
    This function deliberately does not use the module's CommandTable to ensure it can run without module dependency.

    .LINK
    https://psappdeploytoolkit.com

    #>

    [CmdletBinding(DefaultParameterSetName = 'None')]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Tasks')]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Management.Infrastructure.CimInstance[]]$Tasks,

        [Parameter(Mandatory = $true, ParameterSetName = 'TaskName')]
        [ValidateNotNullOrEmpty()]
        [System.String]$TaskName
    )

    # Remove Debugger values to unblock processes.
    Get-ItemProperty -Path "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\*" -Name Debugger -Verbose:$false -ErrorAction Ignore | & {
        process
        {
            if ($_.Debugger.Contains('Show-ADTBlockedAppDialog'))
            {
                Write-Verbose -Message "Removing the Image File Execution Options registry key to unblock execution of [$($_.PSChildName)]."
                Remove-ItemProperty -LiteralPath $_.PSPath -Name Debugger -Verbose:$false
            }
        }
    }

    # Remove the scheduled task if it exists.
    switch ($PSCmdlet.ParameterSetName)
    {
        TaskName
        {
            Write-Verbose -Message "Deleting Scheduled Task [$TaskName]."
            Get-ScheduledTask -TaskName $TaskName -Verbose:$false -ErrorAction Ignore | Unregister-ScheduledTask -Confirm:$false -Verbose:$false
            break
        }
        Tasks
        {
            Write-Verbose -Message "Deleting Scheduled Tasks ['$($Tasks.TaskName -join "', '")']."
            $Tasks | Unregister-ScheduledTask -Confirm:$false -Verbose:$false
            break
        }
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Write-ADTLogEntryToInformationStream
#
#-----------------------------------------------------------------------------

function Write-ADTLogEntryToInformationStream
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Message,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Source,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Format,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.ConsoleColor]$ForegroundColor,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.ConsoleColor]$BackgroundColor
    )

    begin
    {
        # Remove parameters that aren't used to generate an InformationRecord object.
        $null = $PSBoundParameters.Remove('Source')
        $null = $PSBoundParameters.Remove('Format')

        # Establish the base InformationRecord to write out.
        $infoRecord = [System.Management.Automation.InformationRecord]::new([System.Management.Automation.HostInformationMessage]$PSBoundParameters, $Source)
    }

    process
    {
        # Update the message for piped operations and write out to the InformationStream.
        $infoRecord.MessageData.Message = [System.String]::Format($Format, $Message)
        $PSCmdlet.WriteInformation($infoRecord)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Add-ADTEdgeExtension
#
#-----------------------------------------------------------------------------

function Add-ADTEdgeExtension
{
    <#
    .SYNOPSIS
        Adds an extension for Microsoft Edge using the ExtensionSettings policy.

    .DESCRIPTION
        This function adds an extension for Microsoft Edge using the ExtensionSettings policy: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-manage-extensions-ref-guide.

        This enables Edge Extensions to be installed and managed like applications, enabling extensions to be pushed to specific devices or users alongside existing GPO/Intune extension policies.

        This should not be used in conjunction with Edge Management Service which leverages the same registry key to configure Edge extensions.

    .PARAMETER ExtensionID
        The ID of the extension to add.

    .PARAMETER UpdateUrl
        The update URL of the extension. This is the URL where the extension will check for updates.

    .PARAMETER InstallationMode
        The installation mode of the extension. Allowed values: blocked, allowed, removed, force_installed, normal_installed.

    .PARAMETER MinimumVersionRequired
        The minimum version of the extension required for installation.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Add-ADTEdgeExtension -ExtensionID "extensionID" -InstallationMode "force_installed" -UpdateUrl "https://edge.microsoft.com/extensionwebstorebase/v1/crx"

        This example adds the specified extension to be force installed in Microsoft Edge.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ExtensionID,

        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (![System.Uri]::IsWellFormedUriString($_, [System.UriKind]::Absolute))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName UpdateUrl -ProvidedValue $_ -ExceptionMessage 'The specified input is not a valid URL.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$UpdateUrl,

        [Parameter(Mandatory = $true)]
        [ValidateSet('blocked', 'allowed', 'removed', 'force_installed', 'normal_installed')]
        [System.String]$InstallationMode,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$MinimumVersionRequired
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Adding extension with ID [$ExtensionID] using installation mode [$InstallationMode] and update URL [$UpdateUrl]$(if ($MinimumVersionRequired) {" with minimum version required [$MinimumVersionRequired]"})."
        try
        {
            try
            {
                # Set up the additional extension.
                $additionalExtension = @{
                    installation_mode = $InstallationMode
                    update_url = $UpdateUrl
                }

                # Add in the minimum version if specified.
                if ($MinimumVersionRequired)
                {
                    $additionalExtension.Add('minimum_version_required', $MinimumVersionRequired)
                }

                # Get the current extensions from the registry, add our additional one, then convert the result back to JSON.
                $extensionsSettings = & $Script:CommandTable.'Get-ADTEdgeExtensions' |
                    & $Script:CommandTable.'Add-Member' -Name $ExtensionID -Value $additionalExtension -MemberType NoteProperty -Force -PassThru |
                    & $Script:CommandTable.'ConvertTo-Json' -Compress

                # Add the additional extension to the current values, then re-write the definition in the registry.
                $null = & $Script:CommandTable.'Set-ADTRegistryKey' -Key Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge -Name ExtensionSettings -Value $extensionsSettings
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Add-ADTSessionClosingCallback
#
#-----------------------------------------------------------------------------

function Add-ADTSessionClosingCallback
{
    <#
    .SYNOPSIS
        Adds a callback to be executed when the ADT session is closing.

    .DESCRIPTION
        The Add-ADTSessionClosingCallback function registers a callback command to be executed when the ADT session is closing. This function sends the callback to the backend function for processing.

    .PARAMETER Callback
        The callback command(s) to be executed when the ADT session is closing.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Add-ADTSessionClosingCallback -Callback $myCallback

        This example adds the specified callback to be executed when the ADT session is closing.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Closing -Action Add @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Add-ADTSessionFinishingCallback
#
#-----------------------------------------------------------------------------

function Add-ADTSessionFinishingCallback
{
    <#
    .SYNOPSIS
        Adds a callback to be executed when the ADT session is finishing.

    .DESCRIPTION
        The Add-ADTSessionFinishingCallback function registers a callback command to be executed when the ADT session is finishing. This function sends the callback to the backend function for processing.

    .PARAMETER Callback
        The callback command(s) to be executed when the ADT session is finishing.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Add-ADTSessionFinishingCallback -Callback $myCallback

        This example adds the specified callback to be executed when the ADT session is finishing.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Finishing -Action Add @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Add-ADTSessionOpeningCallback
#
#-----------------------------------------------------------------------------

function Add-ADTSessionOpeningCallback
{
    <#
    .SYNOPSIS
        Adds a callback to be executed when the ADT session is opening.

    .DESCRIPTION
        The Add-ADTSessionOpeningCallback function registers a callback command to be executed when the ADT session is opening. This function sends the callback to the backend function for processing.

    .PARAMETER Callback
        The callback command(s) to be executed when the ADT session is opening.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Add-ADTSessionOpeningCallback -Callback $myCallback

        This example adds the specified callback to be executed when the ADT session is opening.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Opening -Action Add @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Add-ADTSessionStartingCallback
#
#-----------------------------------------------------------------------------

function Add-ADTSessionStartingCallback
{
    <#
    .SYNOPSIS
        Adds a callback to be executed when the ADT session is starting.

    .DESCRIPTION
        The Add-ADTSessionStartingCallback function registers a callback command to be executed when the ADT session is starting. This function sends the callback to the backend function for processing.

    .PARAMETER Callback
        The callback command(s) to be executed when the ADT session is starting.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Add-ADTSessionStartingCallback -Callback $myCallback

        This example adds the specified callback to be executed when the ADT session is starting.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Starting -Action Add @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Block-ADTAppExecution
#
#-----------------------------------------------------------------------------

function Block-ADTAppExecution
{
    <#
    .SYNOPSIS
        Block the execution of an application(s).

    .DESCRIPTION
        This function is called when you pass the -BlockExecution parameter to the Stop-RunningApplications function. It does the following:

        1.  Makes a copy of this script in a temporary directory on the local machine.
        2.  Checks for an existing scheduled task from previous failed installation attempt where apps were blocked and if found, calls the Unblock-ADTAppExecution function to restore the original IFEO registry keys.
            This is to prevent the function from overriding the backup of the original IFEO options.
        3.  Creates a scheduled task to restore the IFEO registry key values in case the script is terminated uncleanly by calling the local temporary copy of this script with the parameter -CleanupBlockedApps.
        4.  Modifies the "Image File Execution Options" registry key for the specified process(s) to call this script with the parameter -ShowBlockedAppDialog.
        5.  When the script is called with those parameters, it will display a custom message to the user to indicate that execution of the application has been blocked while the installation is in progress.
            The text of this message can be customized in the strings.psd1 file.

    .PARAMETER ProcessName
        Name of the process or processes separated by commas.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Block-ADTAppExecution -ProcessName ('winword','excel')

        This example blocks the execution of Microsoft Word and Excel.

    .NOTES
        An active ADT session is required to use this function.

        It is used when the -BlockExecution parameter is specified with the Show-ADTInstallationWelcome function to block applications.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, HelpMessage = 'Specify process names, separated by commas.')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ProcessName
    )

    begin
    {
        # Get everything we need before commencing.
        try
        {
            $adtSession = & $Script:CommandTable.'Get-ADTSession'
            $adtEnv = & $Script:CommandTable.'Get-ADTEnvironmentTable'
            $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
            $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $taskName = "$($adtEnv.appDeployToolkitName)_$($adtSession.installName)_BlockedApps" -replace $adtEnv.InvalidScheduledTaskNameCharsRegExPattern
    }

    process
    {
        # Bypass if no Admin rights.
        if (!$adtEnv.IsAdmin)
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing Function [$($MyInvocation.MyCommand.Name)], because [User: $($adtEnv.ProcessNTAccount)] is not admin."
            return
        }

        try
        {
            try
            {
                # Clean up any previous state that might be lingering.
                if ($task = & $Script:CommandTable.'Get-ScheduledTask' -TaskName $taskName -ErrorAction Ignore)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Scheduled task [$taskName] already exists, running [Unblock-ADTAppExecution] to clean up previous state."
                    & $Script:CommandTable.'Unblock-ADTAppExecution' -Tasks $task
                }

                # Create a scheduled task to run on startup to call this script and clean up blocked applications in case the installation is interrupted, e.g. user shuts down during installation"
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Creating scheduled task to cleanup blocked applications in case the installation is interrupted.'
                try
                {
                    $nstParams = @{
                        Principal = & $Script:CommandTable.'New-ScheduledTaskPrincipal' -Id Author -UserId S-1-5-18
                        Trigger = & $Script:CommandTable.'New-ScheduledTaskTrigger' -AtStartup
                        Action = & $Script:CommandTable.'New-ScheduledTaskAction' -Execute $adtEnv.envPSProcessPath -Argument "-NonInteractive -NoProfile -NoLogo -WindowStyle Hidden -EncodedCommand $(& $Script:CommandTable.'Out-ADTPowerShellEncodedCommand' -Command "& {$($Script:CommandTable.'Unblock-ADTAppExecutionInternal'.ScriptBlock)} -TaskName '$($taskName.Replace("'", "''"))'")"
                        Settings = & $Script:CommandTable.'New-ScheduledTaskSettingsSet' -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -ExecutionTimeLimit ([System.TimeSpan]::FromHours(1))
                    }
                    $null = & $Script:CommandTable.'New-ScheduledTask' @nstParams | & $Script:CommandTable.'Register-ScheduledTask' -TaskName $taskName
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to create the scheduled task [$taskName]." -Severity 3
                    return
                }

                # Store the BlockExection command in the registry due to IFEO length issues when > 255 chars.
                $blockExecRegPath = & $Script:CommandTable.'Convert-ADTRegistryPath' -Key (& $Script:CommandTable.'Join-Path' -Path $adtConfig.Toolkit.RegPath -ChildPath $adtEnv.appDeployToolkitName)
                $blockExecCommand = "& (Import-Module -FullyQualifiedName @{ ModuleName = '$("$($Script:PSScriptRoot)\$($MyInvocation.MyCommand.Module.Name).psd1".Replace("'", "''"))'; Guid = '$($MyInvocation.MyCommand.Module.Guid)'; ModuleVersion = '$($MyInvocation.MyCommand.Module.Version)' } -PassThru) { & `$CommandTable.'Initialize-ADTModule' -ScriptDirectory '$($Script:ADT.Directories.Script.Replace("'", "''"))'; `$null = & `$CommandTable.'Show-ADTInstallationPrompt$($adtConfig.UI.DialogStyle)' -Title '$($adtSession.InstallTitle.Replace("'","''"))' -Subtitle '$([System.String]::Format($adtStrings.WelcomePrompt.Fluent.Subtitle, $adtSession.DeploymentType).Replace("'", "''"))' -Timeout $($adtConfig.UI.DefaultTimeout) -Message '$($adtStrings.BlockExecution.Message.Replace("'", "''"))' -Icon Warning -ButtonRightText OK }"
                & $Script:CommandTable.'Set-ADTRegistryKey' -Key $blockExecRegPath -Name BlockExecutionCommand -Value $blockExecCommand

                # Enumerate each process and set the debugger value to block application execution.
                foreach ($process in ($ProcessName -replace '$', '.exe'))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Setting the Image File Execution Option registry key to block execution of [$process]."
                    & $Script:CommandTable.'Set-ADTRegistryKey' -Key (& $Script:CommandTable.'Join-Path' -Path 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options' -ChildPath $process) -Name Debugger -Value "conhost.exe --headless $([System.IO.Path]::GetFileName($adtEnv.envPSProcessPath)) $(if (!(& $Script:CommandTable.'Test-ADTModuleIsReleaseBuild')) { "-ExecutionPolicy Bypass " })-NonInteractive -NoProfile -NoLogo -Command & ([scriptblock]::Create([Microsoft.Win32.Registry]::GetValue('$($blockExecRegPath -replace '^Microsoft\.PowerShell\.Core\\Registry::')', 'BlockExecutionCommand', `$null)))"
                }

                # Add callback to remove all blocked app executions during the shutdown of the final session.
                & $Script:CommandTable.'Add-ADTSessionFinishingCallback' -Callback $Script:CommandTable.'Unblock-ADTAppExecution'
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Close-ADTInstallationProgress
#
#-----------------------------------------------------------------------------

function Close-ADTInstallationProgress
{
    <#
    .SYNOPSIS
        Closes the dialog created by Show-ADTInstallationProgress.

    .DESCRIPTION
        Closes the dialog created by Show-ADTInstallationProgress. This function is called by the Close-ADTSession function to close a running instance of the progress dialog if found.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Close-ADTInstallationProgress

        This example closes the dialog created by Show-ADTInstallationProgress.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    begin
    {
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Return early if we're silent, a window wouldn't have ever opened.
                if (!(& $Script:CommandTable.'Test-ADTInstallationProgressRunning'))
                {
                    return
                }
                if ($adtSession -and $adtSession.IsSilent())
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Mode: $($adtSession.DeployMode)]"
                    return
                }

                # Call the underlying function to close the progress window.
                & $Script:CommandTable."$($MyInvocation.MyCommand.Name)$($adtConfig.UI.DialogStyle)"
                & $Script:CommandTable.'Remove-ADTSessionFinishingCallback' -Callback $MyInvocation.MyCommand

                # We only send balloon tips when a session is active.
                if (!$adtSession)
                {
                    return
                }

                # Send out the final toast notification.
                switch ($adtSession.GetDeploymentStatus())
                {
                    ([PSADT.Module.DeploymentStatus]::FastRetry)
                    {
                        & $Script:CommandTable.'Show-ADTBalloonTip' -BalloonTipIcon Warning -BalloonTipText "$($adtSession.GetDeploymentTypeName()) $((& $Script:CommandTable.'Get-ADTStringTable').BalloonText.($_.ToString()))"
                        break
                    }
                    ([PSADT.Module.DeploymentStatus]::Error)
                    {
                        & $Script:CommandTable.'Show-ADTBalloonTip' -BalloonTipIcon Error -BalloonTipText "$($adtSession.GetDeploymentTypeName()) $((& $Script:CommandTable.'Get-ADTStringTable').BalloonText.($_.ToString()))"
                        break
                    }
                    default
                    {
                        & $Script:CommandTable.'Show-ADTBalloonTip' -BalloonTipIcon Info -BalloonTipText "$($adtSession.GetDeploymentTypeName()) $((& $Script:CommandTable.'Get-ADTStringTable').BalloonText.($_.ToString()))"
                        break
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Close-ADTSession
#
#-----------------------------------------------------------------------------

function Close-ADTSession
{
    <#
    .SYNOPSIS
        Closes the active ADT session.

    .DESCRIPTION
        The Close-ADTSession function closes the active ADT session, updates the session's exit code if provided, invokes all registered callbacks, and cleans up the session state. If this is the last session, it flags the module as uninitialized and exits the process with the last exit code.

    .PARAMETER ExitCode
        The exit code to set for the session.

    .PARAMETER Force
        Forcibly exits PowerShell upon closing of the final session.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Close-ADTSession

        This example closes the active ADT session without setting an exit code.

    .EXAMPLE
        Close-ADTSession -ExitCode 0

        This example closes the active ADT session and sets the exit code to 0.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$ExitCode,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Force
    )

    begin
    {
        # Get the active session and throw if we don't have it.
        try
        {
            $adtSession = & $Script:CommandTable.'Get-ADTSession'
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        # Make this function continue on error and ensure the caller doesn't override ErrorAction.
        $PSBoundParameters.ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        # Change the install phase since we've finished initialising. This should get overwritten shortly.
        $adtSession.InstallPhase = 'Finalization'

        # Update the session's exit code with the provided value.
        if ($PSBoundParameters.ContainsKey('ExitCode') -and (!$adtSession.GetExitCode() -or !$ExitCode.Equals(60001)))
        {
            $adtSession.SetExitCode($ExitCode)
        }

        # Invoke all callbacks and capture all errors.
        $callbackErrors = foreach ($callback in $($Script:ADT.Callbacks.Closing; if ($Script:ADT.Sessions.Count.Equals(1)) { $Script:ADT.Callbacks.Finishing }))
        {
            try
            {
                try
                {
                    & $callback
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                $_; & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failure occurred while invoking callback [$($callback.Name)]."
            }
        }

        # Close out the active session and clean up session state.
        try
        {
            try
            {
                & $Script:CommandTable.'New-Variable' -Name ExitCode -Value $adtSession.Close() -Force -Confirm:$false
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failure occurred while closing ADTSession for [$($adtSession.InstallName)]."
            $ExitCode = 60001
        }

        # Hand over to our backend closure routine if this was the last session.
        if (!$Script:ADT.Sessions.Count)
        {
            & $Script:CommandTable.'Exit-ADTInvocation' -ExitCode $ExitCode -Force:($Force -or ($Host.Name.Equals('ConsoleHost') -and $callbackErrors))
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Complete-ADTFunction
#
#-----------------------------------------------------------------------------

function Complete-ADTFunction
{
    <#
    .SYNOPSIS
        Completes the execution of an ADT function.

    .DESCRIPTION
        The Complete-ADTFunction function finalizes the execution of an ADT function by writing a debug log message and restoring the original global verbosity if it was archived off.

    .PARAMETER Cmdlet
        The PSCmdlet object representing the cmdlet being completed.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Complete-ADTFunction -Cmdlet $PSCmdlet

        This example completes the execution of the current ADT function.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )

    # Write debug log messages and restore original global verbosity if a value was archived off.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Function End' -Source $Cmdlet.MyInvocation.MyCommand.Name -DebugMessage
}


#-----------------------------------------------------------------------------
#
# MARK: Convert-ADTRegistryPath
#
#-----------------------------------------------------------------------------

function Convert-ADTRegistryPath
{
    <#
    .SYNOPSIS
        Converts the specified registry key path to a format that is compatible with built-in PowerShell cmdlets.

    .DESCRIPTION
        Converts the specified registry key path to a format that is compatible with built-in PowerShell cmdlets.

        Converts registry key hives to their full paths. Example: HKLM is converted to "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE".

    .PARAMETER Key
        Path to the registry key to convert (can be a registry hive or fully qualified path)

    .PARAMETER Wow6432Node
        Specifies that the 32-bit registry view (Wow6432Node) should be used on a 64-bit system.

    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.

        Specify this parameter from the Invoke-ADTAllUsersRegistryAction function to read/edit HKCU registry settings for all users on the system.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the converted registry key path.

    .EXAMPLE
        Convert-ADTRegistryPath -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1AD147D0-BE0E-3D6C-AC11-64F6DC4163F1}'

        Converts the specified registry key path to a format compatible with PowerShell cmdlets.

    .EXAMPLE
        Convert-ADTRegistryPath -Key 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1AD147D0-BE0E-3D6C-AC11-64F6DC4163F1}'

        Converts the specified registry key path to a format compatible with PowerShell cmdlets.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Suppress logging output unless the caller has said otherwise.
        if (!$PSBoundParameters.ContainsKey('InformationAction'))
        {
            $InformationPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
        }
    }

    process
    {
        try
        {
            try
            {
                # Convert the registry key hive to the full path, only match if at the beginning of the line.
                $Script:Registry.PathReplacements.GetEnumerator() | . {
                    process
                    {
                        if ($Key -match $_.Key)
                        {
                            foreach ($regexMatch in ($Script:Registry.PathMatches -replace '^', $_.Key))
                            {
                                $Key = $Key -replace $regexMatch, $_.Value
                            }
                        }
                    }
                }

                # Process the WOW6432Node values if applicable.
                if ($Wow6432Node -and [System.Environment]::Is64BitProcess)
                {
                    $Script:Registry.WOW64Replacements.GetEnumerator() | . {
                        process
                        {
                            if ($Key -match $_.Key)
                            {
                                $Key = $Key -replace $_.Key, $_.Value
                            }
                        }
                    }
                }

                # Append the PowerShell provider to the registry key path.
                if ($Key -notmatch '^Microsoft\.PowerShell\.Core\\Registry::')
                {
                    $Key = "Microsoft.PowerShell.Core\Registry::$key"
                }

                # If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID.
                if ($PSBoundParameters.ContainsKey('SID'))
                {
                    if ($Key -match '^Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER\\')
                    {
                        $Key = $Key -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER\\', "Microsoft.PowerShell.Core\Registry::HKEY_USERS\$SID\"
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'SID parameter specified but the registry hive of the key is not HKEY_CURRENT_USER.' -Severity 2
                        return
                    }
                }

                # Check for expected key string format.
                if ($Key -notmatch '^Microsoft\.PowerShell\.Core\\Registry::HKEY_(LOCAL_MACHINE|CLASSES_ROOT|CURRENT_USER|USERS|CURRENT_CONFIG|PERFORMANCE_DATA)')
                {
                    $naerParams = @{
                        Exception = [System.ArgumentException]::new("Unable to detect target registry hive in string [$Key].")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                        ErrorId = 'RegistryKeyValueInvalid'
                        TargetObject = $Key
                        RecommendedAction = "Please confirm the supplied value is correct and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Return fully qualified registry key path [$Key]."
                return $Key
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Convert-ADTValuesFromRemainingArguments
#
#-----------------------------------------------------------------------------

function Convert-ADTValuesFromRemainingArguments
{
    <#
    .SYNOPSIS
        Converts the collected values from a ValueFromRemainingArguments parameter value into a dictionary or PowerShell.exe command line arguments.

    .DESCRIPTION
        This function converts the collected values from a ValueFromRemainingArguments parameter value into a dictionary or PowerShell.exe command line arguments.

    .PARAMETER RemainingArguments
        The collected values to enumerate and process into a dictionary.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.Generic.Dictionary[System.String, System.Object]

        Convert-ADTValuesFromRemainingArguments returns a dictionary of the processed input.

    .EXAMPLE
        Convert-ADTValuesFromRemainingArguments -RemainingArguments $args

        Converts an $args array into a $PSBoundParameters-compatible dictionary.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[System.String, System.Object]])]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowNull()][AllowEmptyCollection()]
        [System.Collections.Generic.List[System.Object]]$RemainingArguments
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Process input into a dictionary and return it. Assume anything starting with a '-' is a new variable.
                return [PSADT.Shared.Utility]::ConvertValuesFromRemainingArguments($RemainingArguments)
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Convert-ADTValueType
#
#-----------------------------------------------------------------------------

function Convert-ADTValueType
{
    <#
    .SYNOPSIS
        Casts the provided value to the requested type without range errors.

    .DESCRIPTION
        This function uses C# code to cast the provided value to the requested type. This avoids errors from PowerShell when values exceed the casted value type's range.

    .PARAMETER Value
        The value to convert.

    .PARAMETER To
        What to cast the value to.

    .INPUTS
        System.Int64

        Convert-ADTValueType will accept any value type as a signed 64-bit integer, then cast to the requested type.

    .OUTPUTS
        System.ValueType

        Convert-ADTValueType will convert the piped input to this type if specified by the caller.

    .EXAMPLE
        Convert-ADTValueType -Value 256 -To SByte

        Invokes the Convert-ADTValueType function and returns the value as a byte, which would equal 0.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Int64]$Value,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Shared.ValueTypes]$To
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $method = "To$To"
    }

    process
    {
        try
        {
            try
            {
                # Use our custom converter to get it done.
                return [PSADT.Shared.ValueTypeConverter]::$method($Value)
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: ConvertTo-ADTNTAccountOrSID
#
#-----------------------------------------------------------------------------

function ConvertTo-ADTNTAccountOrSID
{
    <#

    .SYNOPSIS
        Convert between NT Account names and their security identifiers (SIDs).

    .DESCRIPTION
        Specify either the NT Account name or the SID and get the other. Can also convert well known sid types.

    .PARAMETER AccountName
        The Windows NT Account name specified in <domain>\<username> format.

        Use fully qualified account names (e.g., <domain>\<username>) instead of isolated names (e.g, <username>) because they are unambiguous and provide better performance.

    .PARAMETER SID
        The Windows NT Account SID.

    .PARAMETER WellKnownSIDName
        Specify the Well Known SID name translate to the actual SID (e.g., LocalServiceSid).

        To get all well known SIDs available on system: [Enum]::GetNames([Security.Principal.WellKnownSidType])

    .PARAMETER WellKnownToNTAccount
        Convert the Well Known SID to an NTAccount name.

    .PARAMETER LocalHost
        Avoids a costly domain check when only converting local accounts.

    .INPUTS
        System.String

        Accepts a string containing the NT Account name or SID.

    .OUTPUTS
        System.String

        Returns the NT Account name or SID.

    .EXAMPLE
        ConvertTo-ADTNTAccountOrSID -AccountName 'CONTOSO\User1'

        Converts a Windows NT Account name to the corresponding SID.

    .EXAMPLE
        ConvertTo-ADTNTAccountOrSID -SID 'S-1-5-21-1220945662-2111687655-725345543-14012660'

        Converts a Windows NT Account SID to the corresponding NT Account Name.

    .EXAMPLE
        ConvertTo-ADTNTAccountOrSID -WellKnownSIDName 'NetworkServiceSid'

        Converts a Well Known SID name to a SID.

    .NOTES
        An active ADT session is NOT required to use this function.

        The conversion can return an empty result if the user account does not exist anymore or if translation fails Refer to: http://blogs.technet.com/b/askds/archive/2011/07/28/troubleshooting-sid-translation-failures-from-the-obvious-to-the-not-so-obvious.aspx

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com

    .LINK
        http://msdn.microsoft.com/en-us/library/system.security.principal.wellknownsidtype(v=vs.110).aspx

    #>

    [CmdletBinding()]
    [OutputType([System.Security.Principal.SecurityIdentifier])]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'NTAccountToSID', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Security.Principal.NTAccount]$AccountName,

        [Parameter(Mandatory = $true, ParameterSetName = 'SIDToNTAccount', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Security.Principal.SecurityIdentifier]$SID,

        [Parameter(Mandatory = $true, ParameterSetName = 'WellKnownName', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Security.Principal.WellKnownSidType]$WellKnownSIDName,

        [Parameter(Mandatory = $false, ParameterSetName = 'WellKnownName')]
        [System.Management.Automation.SwitchParameter]$WellKnownToNTAccount,

        [Parameter(Mandatory = $false, ParameterSetName = 'WellKnownName')]
        [System.Management.Automation.SwitchParameter]$LocalHost
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            SIDToNTAccount
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Converting $(($msg = "the SID [$SID] to an NT Account name"))."
                try
                {
                    return $SID.Translate([System.Security.Principal.NTAccount])
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to convert $msg. It may not be a valid account anymore or there is some other problem.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 2
                }
                break
            }
            NTAccountToSID
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Converting $(($msg = "the NT Account [$AccountName] to a SID"))."
                try
                {
                    return $AccountName.Translate([System.Security.Principal.SecurityIdentifier])
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to convert $msg. It may not be a valid account anymore or there is some other problem.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 2
                }
                break
            }
            WellKnownName
            {
                # Get the SID for the root domain.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Converting $(($msg = "the Well Known SID Name [$WellKnownSIDName] to a $(('SID', 'NTAccount')[!!$WellKnownToNTAccount])"))."
                $DomainSid = if (!$LocalHost)
                {
                    try
                    {
                        [System.Security.Principal.SecurityIdentifier]::new([System.DirectoryServices.DirectoryEntry]::new("LDAP://$((& $Script:CommandTable.'Get-CimInstance' -ClassName Win32_ComputerSystem).Domain.ToLower())").ObjectSid[0], 0)
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Unable to get Domain SID from Active Directory. Setting Domain SID to $null.' -Severity 2
                    }
                }

                # Get the SID for the well known SID name.
                try
                {
                    $NTAccountSID = [System.Security.Principal.SecurityIdentifier]::new($WellKnownSIDName, $DomainSid)
                    if ($WellKnownToNTAccount)
                    {
                        return $NTAccountSID.Translate([System.Security.Principal.NTAccount])
                    }
                    return $NTAccountSID
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to convert $msg. It may not be a valid account anymore or there is some other problem.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                }
                break
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Copy-ADTContentToCache
#
#-----------------------------------------------------------------------------

function Copy-ADTContentToCache
{
    <#
    .SYNOPSIS
        Copies the toolkit content to a cache folder on the local machine and sets the $adtSession.DirFiles and $adtSession.DirSupportFiles directory to the cache path.

    .DESCRIPTION
        Copies the toolkit content to a cache folder on the local machine and sets the $adtSession.DirFiles and $adtSession.DirSupportFiles directory to the cache path.

        This function is useful in environments where an Endpoint Management solution does not provide a managed cache for source files, such as Intune.

        It is important to clean up the cache in the uninstall section for the current version and potentially also in the pre-installation section for previous versions.

    .PARAMETER Path
        The path to the software cache folder.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Copy-ADTContentToCache -Path "$envWinDir\Temp\PSAppDeployToolkit"

        This example copies the toolkit content to the specified cache folder.

    .NOTES
        An active ADT session is required to use this function.

        This can be used in the absence of an Endpoint Management solution that provides a managed cache for source files, e.g. Intune is lacking this functionality whereas ConfigMgr includes this functionality.

        Since this cache folder is effectively unmanaged, it is important to cleanup the cache in the uninstall section for the current version and potentially also in the pre-installation section for previous versions.

        This can be done using `Remove-ADTFile -Path "(Get-ADTConfig).Toolkit.CachePath\$($adtSession.InstallName)" -Recurse -ErrorAction Ignore`.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path = "$((& $Script:CommandTable.'Get-ADTConfig').Toolkit.CachePath)\$((& $Script:CommandTable.'Get-ADTSession').installName)"
    )

    begin
    {
        try
        {
            $adtSession = & $Script:CommandTable.'Get-ADTSession'
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        # Create the cache folder if it does not exist.
        if (![System.IO.Directory]::Exists($Path))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating cache folder [$Path]."
            try
            {
                try
                {
                    $null = & $Script:CommandTable.'New-Item' -Path $Path -ItemType Directory
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to create cache folder [$Path]."
                return
            }
        }
        else
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Cache folder [$Path] already exists."
        }

        # Copy the toolkit content to the cache folder.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying toolkit content to cache folder [$Path]."
        try
        {
            try
            {
                & $Script:CommandTable.'Copy-ADTFile' -Path (& $Script:CommandTable.'Join-Path' $adtSession.ScriptDirectory '*') -Destination $Path -Recurse
                $adtSession.DirFiles = "$Path\Files"
                $adtSession.DirSupportFiles = "$Path\SupportFiles"
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to copy toolkit content to cache folder [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Copy-ADTFile
#
#-----------------------------------------------------------------------------

function Copy-ADTFile
{
    <#
    .SYNOPSIS
        Copies files and directories from a source to a destination.

    .DESCRIPTION
        Copies files and directories from a source to a destination. This function supports recursive copying, overwriting existing files, and returning the copied items.

    .PARAMETER Path
        Path of the file to copy. Multiple paths can be specified.

    .PARAMETER Destination
        Destination Path of the file to copy.

    .PARAMETER Recurse
        Copy files in subdirectories.

    .PARAMETER Flatten
        Flattens the files into the root destination directory.

    .PARAMETER ContinueFileCopyOnError
        Continue copying files if an error is encountered. This will continue the deployment script and will warn about files that failed to be copied.

    .PARAMETER FileCopyMode
        Select from 'Native' or 'Robocopy'. Default is configured in config.psd1. Note that Robocopy supports * in file names, but not folders, in source paths.

    .PARAMETER RobocopyParams
        Override the default Robocopy parameters. Default is: /NJH /NJS /NS /NC /NP /NDL /FP /IS /IT /IM /XX /MT:4 /R:1 /W:1

    .PARAMETER RobocopyAdditionalParams
        Append to the default Robocopy parameters. Default is: /NJH /NJS /NS /NC /NP /NDL /FP /IS /IT /IM /XX /MT:4 /R:1 /W:1

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Copy-ADTFile -Path 'C:\Path\file.txt' -Destination 'D:\Destination\file.txt'

        Copies the file 'file.txt' from 'C:\Path' to 'D:\Destination'.

    .EXAMPLE
        Copy-ADTFile -Path 'C:\Path\Folder' -Destination 'D:\Destination\Folder' -Recurse

        Recursively copies the folder 'Folder' from 'C:\Path' to 'D:\Destination'.

    .EXAMPLE
        Copy-ADTFile -Path 'C:\Path\file.txt' -Destination 'D:\Destination\file.txt'

        Copies the file 'file.txt' from 'C:\Path' to 'D:\Destination', overwriting the destination file if it exists.

    .EXAMPLE
        Copy-ADTFile -Path "$($adtSession.DirFiles)\*" -Destination C:\some\random\file\path

        Copies all files within the active session's Files folder to 'C:\some\random\file\path', overwriting the destination file if it exists.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

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
        [System.Management.Automation.SwitchParameter]$Recurse = $false,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Flatten,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ContinueFileCopyOnError,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Native', 'Robocopy')]
        [System.String]$FileCopyMode,

        [Parameter(Mandatory = $false)]
        [System.String]$RobocopyParams = '/NJH /NJS /NS /NC /NP /NDL /FP /IS /IT /IM /XX /MT:4 /R:1 /W:1',

        [Parameter(Mandatory = $false)]
        [System.String]$RobocopyAdditionalParams

    )

    begin
    {
        # If a FileCopyMode hasn't been specified, potentially initialize the module so we can get it from the config.
        if (!$PSBoundParameters.ContainsKey('FileCopyMode'))
        {
            $null = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
            $FileCopyMode = (& $Script:CommandTable.'Get-ADTConfig').Toolkit.FileCopyMode
        }

        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Verify that Robocopy can be used if selected
        if ($FileCopyMode -eq 'Robocopy')
        {
            # Check if Robocopy is on the system.
            if (& $Script:CommandTable.'Test-Path' -Path "$([System.Environment]::SystemDirectory)\Robocopy.exe" -PathType Leaf)
            {
                # Disable Robocopy if $Path has a folder containing a * wildcard.
                if ($Path -match '\*.*\\')
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Asterisk wildcard specified in folder portion of path variable. Falling back to native PowerShell method." -Severity 2
                    $FileCopyMode = 'Native'
                }
                # Don't just check for an extension here, also check for base name without extension to allow copying to a directory such as .config.
                elseif ([System.IO.Path]::HasExtension($Destination) -and [System.IO.Path]::GetFileNameWithoutExtension($Destination) -and !(& $Script:CommandTable.'Test-Path' -LiteralPath $Destination -PathType Container))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Destination path appears to be a file. Falling back to native PowerShell method." -Severity 2
                    $FileCopyMode = 'Native'
                }
                else
                {
                    $robocopyCommand = "$([System.Environment]::SystemDirectory)\Robocopy.exe"

                    if ($Recurse -and !$Flatten)
                    {
                        # Add /E to Robocopy parameters if it is not already included.
                        if ($RobocopyParams -notmatch '/E(\s+|$)' -and $RobocopyAdditionalParams -notmatch '/E(\s+|$)')
                        {
                            $RobocopyParams = $RobocopyParams + " /E"
                        }
                    }
                    else
                    {
                        # Ensure that /E is not included in the Robocopy parameters as it will copy recursive folders.
                        $RobocopyParams = $RobocopyParams -replace '/E(\s+|$)'
                        $RobocopyAdditionalParams = $RobocopyAdditionalParams -replace '/E(\s+|$)'
                    }

                    # Older versions of Robocopy do not support /IM, remove if unsupported.
                    if ((& $robocopyCommand /?) -notmatch '/IM\s')
                    {
                        $RobocopyParams = $RobocopyParams -replace '/IM(\s+|$)'
                        $RobocopyAdditionalParams = $RobocopyAdditionalParams -replace '/IM(\s+|$)'
                    }
                }
            }
            else
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy is not available on this system. Falling back to native PowerShell method." -Severity 2
                $FileCopyMode = 'Native'
            }
        }
    }

    process
    {
        if ($FileCopyMode -eq 'Robocopy')
        {
            foreach ($srcPath in $Path)
            {
                try
                {
                    if (!(& $Script:CommandTable.'Test-Path' -Path $srcPath))
                    {
                        if (!$ContinueFileCopyOnError)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Source path [$srcPath] not found." -Severity 2
                            $naerParams = @{
                                Exception = [System.IO.FileNotFoundException]::new("Source path [$srcPath] not found.")
                                Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                                ErrorId = 'FileNotFoundError'
                                TargetObject = $srcPath
                                RecommendedAction = 'Please verify that the path is accessible and try again.'
                            }
                            & $Script:CommandTable.'Write-Error' -ErrorRecord (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                        }
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Source path [$srcPath] not found. Will continue due to ContinueFileCopyOnError = `$true." -Severity 2
                        continue
                    }

                    # Pre-create destination folder if it does not exist; Robocopy will auto-create non-existent destination folders, but pre-creating ensures we can use Resolve-Path.
                    if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $Destination -PathType Container))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Destination assumed to be a folder which does not exist, creating destination folder [$Destination]."
                        $null = & $Script:CommandTable.'New-Item' -Path $Destination -Type Directory -Force
                    }

                    # If source exists as a folder, append the last subfolder to the destination, so that Robocopy produces similar results to native PowerShell.
                    if (& $Script:CommandTable.'Test-Path' -LiteralPath $srcPath -PathType Container)
                    {
                        # Trim ending backslash from paths which can cause problems with Robocopy.
                        # Resolve paths in case relative paths beggining with .\, ..\, or \ are used.
                        # Strip Microsoft.PowerShell.Core\FileSystem:: from the beginning of the resulting string, since Resolve-Path adds this to UNC paths.
                        $robocopySource = (& $Script:CommandTable.'Resolve-Path' -LiteralPath $srcPath.TrimEnd('\')).Path -replace '^Microsoft\.PowerShell\.Core\\FileSystem::'
                        $robocopyDestination = & $Script:CommandTable.'Join-Path' ((& $Script:CommandTable.'Resolve-Path' -LiteralPath $Destination).Path -replace '^Microsoft\.PowerShell\.Core\\FileSystem::') (& $Script:CommandTable.'Split-Path' -Path $srcPath -Leaf)
                        $robocopyFile = '*'
                    }
                    else
                    {
                        # Else assume source is a file and split args to the format <SourceFolder> <DestinationFolder> <FileName>.
                        # Trim ending backslash from paths which can cause problems with Robocopy.
                        # Resolve paths in case relative paths beggining with .\, ..\, or \ are used.
                        # Strip Microsoft.PowerShell.Core\FileSystem:: from the beginning of the resulting string, since Resolve-Path adds this to UNC paths.
                        $ParentPath = & $Script:CommandTable.'Split-Path' -Path $srcPath -Parent
                        $robocopySource = if ([System.String]::IsNullOrWhiteSpace($ParentPath))
                        {
                            $ExecutionContext.SessionState.Path.CurrentLocation.Path
                        }
                        else
                        {
                           (& $Script:CommandTable.'Resolve-Path' -LiteralPath $ParentPath).Path -replace '^Microsoft\.PowerShell\.Core\\FileSystem::'
                        }
                        $robocopyDestination = (& $Script:CommandTable.'Resolve-Path' -LiteralPath $Destination.TrimEnd('\')).Path -replace '^Microsoft\.PowerShell\.Core\\FileSystem::'
                        $robocopyFile = (& $Script:CommandTable.'Split-Path' -Path $srcPath -Leaf)
                    }

                    # Set up copy operation.
                    if ($Flatten)
                    {
                        # Copy all files from the root source folder.
                        $copyFileSplat = @{
                            Destination = $Destination  # Use the original destination path, not $robocopyDestination which could have had a subfolder appended to it.
                            Recurse = $false  # Disable recursion as this will create subfolders in the destination.
                            Flatten = $false  # Disable flattening to prevent infinite loops.
                            ContinueFileCopyOnError = $ContinueFileCopyOnError
                            FileCopyMode = $FileCopyMode
                            RobocopyParams = $RobocopyParams
                            RobocopyAdditionalParams = $RobocopyAdditionalParams
                        }
                        if ($PSBoundParameters.ContainsKey('ErrorAction'))
                        {
                            $copyFileSplat.ErrorAction = $PSBoundParameters.ErrorAction
                        }
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying file(s) recursively in path [$srcPath] to destination [$Destination] root folder, flattened."
                        & $Script:CommandTable.'Copy-ADTFile' @copyFileSplat -Path ((& $Script:CommandTable.'Join-Path' $robocopySource $robocopyFile))

                        # Copy all files from subfolders, appending file name to subfolder path and repeat Copy-ADTFile.
                        & $Script:CommandTable.'Get-ChildItem' -Path $robocopySource -Directory -Recurse -Force -ErrorAction Ignore | & {
                            process
                            {
                                & $Script:CommandTable.'Copy-ADTFile' @copyFileSplat -Path (& $Script:CommandTable.'Join-Path' $_.FullName $robocopyFile)
                            }
                        }

                        # Skip to next $srcPath in $Path since we have handed off all copy tasks to separate executions of the function.
                        continue
                    }
                    elseif ($Recurse)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying file(s) recursively in path [$srcPath] to destination [$Destination]."
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying file(s) in path [$srcPath] to destination [$Destination]."
                    }

                    # Create new directory if it doesn't exist.
                    if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $robocopyDestination -PathType Container))
                    {
                        $null = & $Script:CommandTable.'New-Item' -Path $robocopyDestination -Type Directory -Force
                    }

                    # Backup destination folder attributes in case known Robocopy bug overwrites them.
                    $destFolderAttributes = [System.IO.File]::GetAttributes($robocopyDestination)

                    # Begin copy operation.
                    $robocopyArgs = "`"$robocopySource`" `"$robocopyDestination`" `"$robocopyFile`" $RobocopyParams $RobocopyAdditionalParams"
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Executing Robocopy command: $robocopyCommand $robocopyArgs"
                    $robocopyResult = & $Script:CommandTable.'Start-ADTProcess' -FilePath $robocopyCommand -ArgumentList $robocopyArgs -CreateNoWindow -PassThru -SuccessExitCodes 0, 1, 2, 3, 4, 5, 6, 7, 8 -ErrorAction Ignore

                    # Trim the last line plus leading whitespace from each line of Robocopy output.
                    $robocopyOutput = $robocopyResult.StdOut.Trim() -Replace '\n\s+', "`n"
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy output:`n$robocopyOutput"

                    # Restore folder attributes in case Robocopy overwrote them.
                    try
                    {
                        [System.IO.File]::SetAttributes($robocopyDestination, $destFolderAttributes)
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to apply attributes [$destFolderAttributes] destination folder [$robocopyDestination]: $($_.Exception.Message)" -Severity 2
                    }

                    # Process the resulting exit code.
                    switch ($robocopyResult.ExitCode)
                    {
                        0 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. No files were copied. No failure was encountered. No files were mismatched. The files already exist in the destination directory; therefore, the copy operation was skipped."; break }
                        1 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. All files were copied successfully."; break }
                        2 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. There are some additional files in the destination directory that aren't present in the source directory. No files were copied."; break }
                        3 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. Some files were copied. Additional files were present. No failure was encountered."; break }
                        4 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. Some Mismatched files or directories were detected. Examine the output log. Housekeeping might be required." -Severity 2; break }
                        5 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. Some files were copied. Some files were mismatched. No failure was encountered."; break }
                        6 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. Additional files and mismatched files exist. No files were copied and no failures were encountered meaning that the files already exist in the destination directory." -Severity 2; break }
                        7 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. Files were copied, a file mismatch was present, and additional files were present." -Severity 2; break }
                        8 { & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy completed. Several files didn't copy." -Severity 2; break }
                        16
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy error [$($robocopyResult.ExitCode)]: Serious error. Robocopy did not copy any files. Either a usage error or an error due to insufficient access privileges on the source or destination directories." -Severity 3
                            if (!$ContinueFileCopyOnError)
                            {
                                $naerParams = @{
                                    Exception = [System.Management.Automation.ApplicationFailedException]::new("Robocopy error $($robocopyResult.ExitCode): Failed to copy file(s) in path [$srcPath] to destination [$Destination]: $robocopyOutput")
                                    Category = [System.Management.Automation.ErrorCategory]::OperationStopped
                                    ErrorId = 'RobocopyError'
                                    TargetObject = $srcPath
                                    RecommendedAction = "Please verify that Path and Destination are accessible and try again."
                                }
                                & $Script:CommandTable.'Write-Error' -ErrorRecord (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                            }
                            break
                        }
                        default
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Robocopy error [$($robocopyResult.ExitCode)]. Unknown Robocopy error." -Severity 3
                            if (!$ContinueFileCopyOnError)
                            {
                                $naerParams = @{
                                    Exception = [System.Management.Automation.ApplicationFailedException]::new("Robocopy error $($robocopyResult.ExitCode): Failed to copy file(s) in path [$srcPath] to destination [$Destination]: $robocopyOutput")
                                    Category = [System.Management.Automation.ErrorCategory]::OperationStopped
                                    ErrorId = 'RobocopyError'
                                    TargetObject = $srcPath
                                    RecommendedAction = "Please verify that Path and Destination are accessible and try again."
                                }
                                & $Script:CommandTable.'Write-Error' -ErrorRecord (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                            }
                            break
                        }
                    }
                }
                catch
                {
                    & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to copy file(s) in path [$srcPath] to destination [$Destination]."
                    if (!$ContinueFileCopyOnError)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'ContinueFileCopyOnError not specified, exiting function.'
                        return
                    }
                }
            }
        }
        elseif ($FileCopyMode -eq 'Native')
        {
            foreach ($srcPath in $Path)
            {
                try
                {
                    try
                    {
                        # If destination has no extension, or if it has an extension only and no name (e.g. a .config folder) and the destination folder does not exist.
                        if ((![System.IO.Path]::HasExtension($Destination) -or ([System.IO.Path]::HasExtension($Destination) -and ![System.IO.Path]::GetFileNameWithoutExtension($Destination))) -and !(& $Script:CommandTable.'Test-Path' -LiteralPath $Destination -PathType Container))
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Destination assumed to be a folder which does not exist, creating destination folder [$Destination]."
                            $null = & $Script:CommandTable.'New-Item' -Path $Destination -Type Directory -Force
                        }

                        # If destination appears to be a file name but parent folder does not exist, create it.
                        if ([System.IO.Path]::HasExtension($Destination) -and [System.IO.Path]::GetFileNameWithoutExtension($Destination) -and !(& $Script:CommandTable.'Test-Path' -LiteralPath ($destinationParent = & $Script:CommandTable.'Split-Path' $Destination -Parent) -PathType Container))
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Destination assumed to be a file whose parent folder does not exist, creating destination folder [$destinationParent]."
                            $null = & $Script:CommandTable.'New-Item' -Path $destinationParent -Type Directory -Force
                        }

                        # Set up parameters for Copy-Item operation.
                        $ciParams = @{
                            Destination = $Destination
                            Force = $true
                        }
                        if ($ContinueFileCopyOnError)
                        {
                            $ciParams.Add('ErrorAction', [System.Management.Automation.ActionPreference]::SilentlyContinue)
                            $ciParams.Add('ErrorVariable', 'FileCopyError')
                        }

                        # Perform copy operation.
                        $null = if ($Flatten)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying file(s) recursively in path [$srcPath] to destination [$Destination] root folder, flattened."
                            if ($srcPaths = & $Script:CommandTable.'Get-ChildItem' -Path $srcPath -File -Recurse -Force -ErrorAction Ignore)
                            {
                                & $Script:CommandTable.'Copy-Item' -LiteralPath $srcPaths.PSPath @ciParams
                            }
                        }
                        elseif ($Recurse)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying file(s) recursively in path [$srcPath] to destination [$Destination]."
                            & $Script:CommandTable.'Copy-Item' -Path $srcPath -Recurse @ciParams
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying file in path [$srcPath] to destination [$Destination]."
                            & $Script:CommandTable.'Copy-Item' -Path $srcPath @ciParams
                        }

                        # Measure success.
                        if ($ContinueFileCopyOnError -and (& $Script:CommandTable.'Test-Path' -LiteralPath Microsoft.PowerShell.Core\Variable::FileCopyError))
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "The following warnings were detected while copying file(s) in path [$srcPath] to destination [$Destination].`n$FileCopyError" -Severity 2
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'File copy completed successfully.'
                        }
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                    }
                }
                catch
                {
                    & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to copy file(s) in path [$srcPath] to destination [$Destination]."
                    if (!$ContinueFileCopyOnError)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'ContinueFileCopyOnError not specified, exiting function.'
                        return
                    }
                }
            }
        }
    }
    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Copy-ADTFileToUserProfiles
#
#-----------------------------------------------------------------------------

function Copy-ADTFileToUserProfiles
{
    <#
    .SYNOPSIS
        Copy one or more items to each user profile on the system.

    .DESCRIPTION
        The Copy-ADTFileToUserProfiles function copies one or more items to each user profile on the system. It supports various options such as recursion, flattening files, and using Robocopy to overcome the 260 character limit.

    .PARAMETER Path
        The path of the file or folder to copy.

    .PARAMETER Destination
        The path of the destination folder to append to the root of the user profile.

    .PARAMETER BasePath
        The base path to append the destination folder to. Default is: Profile. Options are: Profile, AppData, LocalAppData, Desktop, Documents, StartMenu, Temp, OneDrive, OneDriveCommercial.

    .PARAMETER Recurse
        Copy files in subdirectories.

    .PARAMETER Flatten
        Flattens the files into the root destination directory.

    .PARAMETER ContinueFileCopyOnError
        Continue copying files if an error is encountered. This will continue the deployment script and will warn about files that failed to be copied.

    .PARAMETER FileCopyMode
        Select from 'Native' or 'Robocopy'. Default is configured in config.psd1. Note that Robocopy supports * in file names, but not folders, in source paths.

    .PARAMETER RobocopyParams
        Override the default Robocopy parameters. Default is: /NJH /NJS /NS /NC /NP /NDL /FP /IS /IT /IM /XX /MT:4 /R:1 /W:1

    .PARAMETER RobocopyAdditionalParams
        Append to the default Robocopy parameters. Default is: /NJH /NJS /NS /NC /NP /NDL /FP /IS /IT /IM /XX /MT:4 /R:1 /W:1

    .PARAMETER ExcludeNTAccount
        Specify NT account names in Domain\Username format to exclude from the list of user profiles.

    .PARAMETER IncludeSystemProfiles
        Include system profiles: SYSTEM, LOCAL SERVICE, NETWORK SERVICE. Default is: $false.

    .PARAMETER IncludeServiceProfiles
        Include service profiles where NTAccount begins with NT SERVICE. Default is: $false.

    .PARAMETER ExcludeDefaultUser
        Exclude the Default User. Default is: $false.

    .INPUTS
        System.String[]

        You can pipe in string values for $Path.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Copy-ADTFileToUserProfiles -Path "$($adtSession.DirSupportFiles)\config.txt" -Destination "AppData\Roaming\MyApp"

        Copy a single file to C:\Users\<UserName>\AppData\Roaming\MyApp for each user.

    .EXAMPLE
        Copy-ADTFileToUserProfiles -Path "$($adtSession.DirSupportFiles)\config.txt","$($adtSession.DirSupportFiles)\config2.txt" -Destination "AppData\Roaming\MyApp"

        Copy two files to C:\Users\<UserName>\AppData\Roaming\MyApp for each user.

    .EXAMPLE
        Copy-ADTFileToUserProfiles -Path "$($adtSession.DirFiles)\MyDocs" Destination "MyApp" -BasePath "Documents" -Recurse

        Copy an entire folder recursively to a new MyApp folder under each user's Documents folder.

    .EXAMPLE
        Copy-ADTFileToUserProfiles -Path "$($adtSession.DirFiles)\.appConfigFolder" -Recurse

        Copy an entire folder to C:\Users\<UserName> for each user.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Path,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Destination,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Profile', 'AppData', 'LocalAppData', 'Desktop', 'Documents', 'StartMenu', 'Temp', 'OneDrive', 'OneDriveCommercial')]
        [System.String]$BasePath = 'Profile',

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Flatten,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Native', 'Robocopy')]
        [System.String]$FileCopyMode,

        [Parameter(Mandatory = $false)]
        [System.String]$RobocopyParams,

        [Parameter(Mandatory = $false)]
        [System.String]$RobocopyAdditionalParams,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ExcludeNTAccount,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$IncludeSystemProfiles,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$IncludeServiceProfiles,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$ExcludeDefaultUser,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$ContinueFileCopyOnError
    )

    begin
    {
        # Initalize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Define default params for Copy-ADTFile.
        $CopyFileSplat = @{
            Recurse = $Recurse
            Flatten = $Flatten
            ContinueFileCopyOnError = $ContinueFileCopyOnError
        }
        if ($PSBoundParameters.ContainsKey('FileCopyMode'))
        {
            $CopyFileSplat.FileCopyMode = $PSBoundParameters.FileCopyMode
        }
        if ($PSBoundParameters.ContainsKey('RobocopyParams'))
        {
            $CopyFileSplat.RobocopyParams = $PSBoundParameters.RobocopyParams
        }
        if ($PSBoundParameters.ContainsKey('RobocopyAdditionalParams'))
        {
            $CopyFileSplat.RobocopyAdditionalParams = $PSBoundParameters.RobocopyAdditionalParams
        }
        if ($PSBoundParameters.ContainsKey('ErrorAction'))
        {
            $CopyFileSplat.ErrorAction = $PSBoundParameters.ErrorAction
        }

        # Define default params for Get-ADTUserProfiles.
        $GetUserProfileSplat = @{
            IncludeSystemProfiles = $IncludeSystemProfiles
            IncludeServiceProfiles = $IncludeServiceProfiles
            ExcludeDefaultUser = $ExcludeDefaultUser
        }
        if ($ExcludeNTAccount)
        {
            $GetUserProfileSplat.ExcludeNTAccount = $ExcludeNTAccount
        }
        if ($BasePath -ne 'ProfilePath')
        {
            $GetUserProfileSplat.LoadProfilePaths = $true
        }

        # Collector for all provided paths.
        $sourcePaths = [System.Collections.Specialized.StringCollection]::new()
    }

    process
    {
        # Add all source paths to the collection.
        $sourcePaths.AddRange($Path)
    }

    end
    {
        # Copy all paths to the specified destination.
        foreach ($UserProfile in (& $Script:CommandTable.'Get-ADTUserProfiles' @GetUserProfileSplat))
        {
            if ([string]::IsNullOrWhiteSpace($UserProfile."$BasePath`Path"))
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Skipping user profile [$($UserProfile.NTAccount)] as path [$BasePath`Path] is not available."
                continue
            }
            $dest = & $Script:CommandTable.'Join-Path' $UserProfile."$BasePath`Path" $Destination
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying path [$Path] to $($dest):"
            & $Script:CommandTable.'Copy-ADTFile' -Path $sourcePaths -Destination $dest @CopyFileSplat
        }

        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Disable-ADTTerminalServerInstallMode
#
#-----------------------------------------------------------------------------

function Disable-ADTTerminalServerInstallMode
{
    <#
    .SYNOPSIS
        Changes to user install mode for Remote Desktop Session Host/Citrix servers.

    .DESCRIPTION
        The Disable-ADTTerminalServerInstallMode function changes the server mode to user install mode for Remote Desktop Session Host/Citrix servers. This is useful for ensuring that applications are installed in a way that is compatible with multi-user environments.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Disable-ADTTerminalServerInstallMode

        This example changes the server mode to user install mode for Remote Desktop Session Host/Citrix servers.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        if (!$Script:ADT.TerminalServerMode)
        {
            return
        }

        try
        {
            try
            {
                & $Script:CommandTable.'Invoke-ADTTerminalServerModeChange' -Mode Execute
                $Script:ADT.TerminalServerMode = $false
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Dismount-ADTWimFile
#
#-----------------------------------------------------------------------------

function Dismount-ADTWimFile
{
    <#
    .SYNOPSIS
        Dismounts a WIM file from the specified mount point.

    .DESCRIPTION
        The Dismount-ADTWimFile function dismounts a WIM file from the specified mount point and discards all changes. This function ensures that the specified path is a valid WIM mount point before attempting to dismount.

    .PARAMETER ImagePath
        The path to the WIM file.

    .PARAMETER Path
        The path to the WIM mount point.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Dismount-ADTWimFile -ImagePath 'C:\Path\To\File.wim'

        This example dismounts the WIM file from all its mount points and discards all changes.

    .EXAMPLE
        Dismount-ADTWimFile -Path 'C:\Mount\WIM'

        This example dismounts the WIM file from the specified mount point and discards all changes.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'ImagePath')]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$ImagePath,

        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [System.IO.DirectoryInfo[]]$Path
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        # Loop through all found mounted images.
        foreach ($wimFile in (& $Script:CommandTable.'Get-ADTMountedWimFile' @PSBoundParameters))
        {
            # Announce commencement.
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Dismounting WIM file at path [$($wimFile.Path)]."
            try
            {
                try
                {
                    # Perform the dismount and discard all changes.
                    try
                    {
                        $null = & $Script:CommandTable.'Invoke-ADTCommandWithRetries' -Command $Script:CommandTable.'Dismount-WindowsImage' -Path $wimFile.Path -Discard
                    }
                    catch
                    {
                        # Re-throw if this error is anything other than a file-locked error.
                        if (!$_.Exception.ErrorCode.Equals(-1052638953))
                        {
                            throw
                        }

                        # Get all open file handles for our path.
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "The directory could not be completely unmounted. Checking for any open file handles that can be closed."
                        $exeHandle = "$Script:PSScriptRoot\bin\$([PSADT.OperatingSystem.OSHelper]::GetArchitecture())\handle\handle.exe"
                        $pathRegex = "^$([System.Text.RegularExpressions.Regex]::Escape($($wimFile.Path)))"
                        $pathHandles = & $Script:CommandTable.'Get-ADTProcessHandles' | & { process { if ($_.Name -match $pathRegex) { return $_ } } }

                        # Throw if we have no handles to close, it means we don't know why the WIM didn't dismount.
                        if (!$pathHandles)
                        {
                            throw
                        }

                        # Close all open file handles.
                        foreach ($handle in $pathHandles)
                        {
                            # Close handle using handle.exe. An exit code of 0 is considered successful.
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(($msg = "Closing handle [$($handle.Handle)] for process [$($handle.Process) ($($handle.PID))]"))."
                            $handleResult = & $exeHandle -accepteula -nobanner -c $handle.Handle -p $handle.PID -y
                            if ($Global:LASTEXITCODE.Equals(0))
                            {
                                continue
                            }

                            # If we're here, we had a bad exit code.
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message ($msg = "$msg failed with exit code [$Global:LASTEXITCODE]: $handleResult") -Severity 3
                            $naerParams = @{
                                Exception = [System.Runtime.InteropServices.ExternalException]::new($msg, $Global:LASTEXITCODE)
                                Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                                ErrorId = 'HandleClosureFailure'
                                TargetObject = $handleResult
                                RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                            }
                            throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                        }

                        # Attempt the dismount again.
                        $null = & $Script:CommandTable.'Invoke-ADTCommandWithRetries' -Command $Script:CommandTable.'Dismount-WindowsImage' -Path $wimFile.Path -Discard
                    }
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Successfully dismounted WIM file."
                    & $Script:CommandTable.'Remove-Item' -LiteralPath $wimFile.Path -Force -Confirm:$false
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage 'Error occurred while attempting to dismount WIM file.' -ErrorAction SilentlyContinue
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Enable-ADTTerminalServerInstallMode
#
#-----------------------------------------------------------------------------

function Enable-ADTTerminalServerInstallMode
{
    <#
    .SYNOPSIS
        Changes to user install mode for Remote Desktop Session Host/Citrix servers.

    .DESCRIPTION
        The Enable-ADTTerminalServerInstallMode function changes the server mode to user install mode for Remote Desktop Session Host/Citrix servers. This is useful for ensuring that applications are installed in a way that is compatible with multi-user environments.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Enable-ADTTerminalServerInstallMode

        This example changes the server mode to user install mode for Remote Desktop Session Host/Citrix servers.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        if ($Script:ADT.TerminalServerMode)
        {
            return
        }

        try
        {
            try
            {
                & $Script:CommandTable.'Invoke-ADTTerminalServerModeChange' -Mode Install
                $Script:ADT.TerminalServerMode = $true
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Export-ADTEnvironmentTableToSessionState
#
#-----------------------------------------------------------------------------

function Export-ADTEnvironmentTableToSessionState
{
    <#
    .SYNOPSIS
        Exports the content of `Get-ADTEnvironmentTable` to the provided SessionState as variables.

    .DESCRIPTION
        This function exports the content of `Get-ADTEnvironmentTable` to the provided SessionState as variables.

    .PARAMETER SessionState
        Caller's SessionState.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Export-ADTEnvironmentTableToSessionState -SessionState $ExecutionContext.SessionState

        Invokes the Export-ADTEnvironmentTableToSessionState function and exports the module's environment table to the provided SessionState.

    .EXAMPLE
        Export-ADTEnvironmentTableToSessionState -SessionState $PSCmdlet.SessionState

        Invokes the Export-ADTEnvironmentTableToSessionState function and exports the module's environment table to the provided SessionState.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SessionState]$SessionState
    )

    begin
    {
        # Initialize function and store the environment table on the stack.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        try
        {
            $adtEnv = & $Script:CommandTable.'Get-ADTEnvironmentTable'
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    process
    {
        try
        {
            try
            {
                $null = $ExecutionContext.InvokeCommand.InvokeScript($SessionState, { $args[1].GetEnumerator() | . { process { & $args[0] -Name $_.Key -Value $_.Value -Option ReadOnly -Force } } $args[0] }.Ast.GetScriptBlock(), $Script:CommandTable.'New-Variable', $adtEnv)
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTApplication
#
#-----------------------------------------------------------------------------

function Get-ADTApplication
{
    <#
    .SYNOPSIS
        Retrieves information about installed applications.

    .DESCRIPTION
        Retrieves information about installed applications by querying the registry. You can specify an application name, a product code, or both. Returns information about application publisher, name & version, product code, uninstall string, install source, location, date, and application architecture.

    .PARAMETER Name
        The name of the application to retrieve information for. Performs a contains match on the application display name by default.

    .PARAMETER NameMatch
        Specifies the type of match to perform on the application name. Valid values are 'Contains', 'Exact', 'Wildcard', and 'Regex'. The default value is 'Contains'.

    .PARAMETER ProductCode
        The product code of the application to retrieve information for.

    .PARAMETER ApplicationType
        Specifies the type of application to remove. Valid values are 'All', 'MSI', and 'EXE'. The default value is 'All'.

    .PARAMETER IncludeUpdatesAndHotfixes
        Include matches against updates and hotfixes in results.

    .PARAMETER FilterScript
        A script used to filter the results as they're processed.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.InstalledApplication

        Returns a custom type with information about an installed application:
        - UninstallKey
        - UninstallParentKey
        - UninstallSubKey
        - ProductCode
        - DisplayName
        - DisplayVersion
        - UninstallString
        - QuietUninstallString
        - InstallSource
        - InstallLocation
        - InstallDate
        - Publisher
        - SystemComponent
        - WindowsInstaller
        - Is64BitApplication

    .EXAMPLE
        Get-ADTApplication

        This example retrieves information about all installed applications.

    .EXAMPLE
        Get-ADTApplication -Name 'Acrobat'

        Returns all applications that contain the name 'Acrobat' in the DisplayName.

    .EXAMPLE
        Get-ADTApplication -Name 'Adobe Acrobat Reader' -NameMatch 'Exact'

        Returns all applications that match the name 'Adobe Acrobat Reader' exactly.

    .EXAMPLE
        Get-ADTApplication -ProductCode '{AC76BA86-7AD7-1033-7B44-AC0F074E4100}'

        Returns the application with the specified ProductCode.

    .EXAMPLE
        Get-ADTApplication -Name 'Acrobat' -ApplicationType 'MSI' -FilterScript { $_.Publisher -match 'Adobe' }

        Returns all MSI applications that contain the name 'Acrobat' in the DisplayName and 'Adobe' in the Publisher name.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ProductCode', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ApplicationType', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    [OutputType([PSADT.Types.InstalledApplication])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Name,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Contains', 'Exact', 'Wildcard', 'Regex')]
        [System.String]$NameMatch = 'Contains',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Guid[]]$ProductCode,

        [Parameter(Mandatory = $false)]
        [ValidateSet('All', 'MSI', 'EXE')]
        [System.String]$ApplicationType = 'All',

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeUpdatesAndHotfixes,

        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]$FilterScript
    )

    begin
    {
        # Announce start.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $updatesSkippedCounter = 0
        $uninstallKeyPaths = $(
            'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
            if ([System.Environment]::Is64BitProcess)
            {
                'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
            }
        )

        # If we're filtering by name, set up the relevant FilterScript.
        $nameFilterScript = if ($Name)
        {
            switch ($NameMatch)
            {
                Contains
                {
                    { foreach ($eachName in $Name) { if ($_.DisplayName -like "*$eachName*") { $true; break } } }
                    break
                }
                Exact
                {
                    { foreach ($eachName in $Name) { if ($_.DisplayName -eq $eachName) { $true; break } } }
                    break
                }
                Wildcard
                {
                    { foreach ($eachName in $Name) { if ($_.DisplayName -like $eachName) { $true; break } } }
                    break
                }
                Regex
                {
                    { foreach ($eachName in $Name) { if ($_.DisplayName -match $eachName) { $true; break } } }
                    break
                }
            }
        }
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Getting information for installed applications$(if ($FilterScript) {' matching the provided FilterScript'})..."
        try
        {
            try
            {
                # Create a custom object with the desired properties for the installed applications and sanitize property details.
                $installedApplication = & $Script:CommandTable.'Get-ItemProperty' -Path $uninstallKeyPaths -ErrorAction Ignore | & {
                    process
                    {
                        # Exclude anything without a DisplayName field.
                        if (!$_.PSObject.Properties.Name.Contains('DisplayName') -or [System.String]::IsNullOrWhiteSpace($_.DisplayName))
                        {
                            return
                        }

                        # Bypass any updates or hotfixes.
                        if (!$IncludeUpdatesAndHotfixes -and ($_.DisplayName -match '((?i)kb\d+|(Cumulative|Security) Update|Hotfix)'))
                        {
                            $updatesSkippedCounter++
                            return
                        }

                        # Apply application type filter if specified.
                        $windowsInstaller = !!($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty WindowsInstaller -ErrorAction Ignore)
                        if ((($ApplicationType -eq 'MSI') -and !$windowsInstaller) -or (($ApplicationType -eq 'EXE') -and $windowsInstaller))
                        {
                            return
                        }

                        # Apply ProductCode filter if specified.
                        $defaultGuid = [System.Guid]::Empty
                        $appMsiGuid = if ($windowsInstaller -and [System.Guid]::TryParse($_.PSChildName, [ref]$defaultGuid)) { $defaultGuid }
                        if ($ProductCode -and (!$appMsiGuid -or ($ProductCode -notcontains $appMsiGuid)))
                        {
                            return
                        }

                        # Apply name filter if specified.
                        if ($nameFilterScript -and !(& $nameFilterScript))
                        {
                            return
                        }

                        # Build out the app object here before we filter as the caller needs to be able to filter on the object's properties.
                        $app = [PSADT.Types.InstalledApplication]::new(
                            $_.PSPath,
                            $_.PSParentPath,
                            $_.PSChildName,
                            $appMsiGuid,
                            $_.DisplayName,
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty DisplayVersion -ErrorAction Ignore),
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty UninstallString -ErrorAction Ignore),
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty QuietUninstallString -ErrorAction Ignore),
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty InstallSource -ErrorAction Ignore),
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty InstallLocation -ErrorAction Ignore),
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty InstallDate -ErrorAction Ignore),
                            ($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty Publisher -ErrorAction Ignore),
                            !!($_ | & $Script:CommandTable.'Select-Object' -ExpandProperty SystemComponent -ErrorAction Ignore),
                            $windowsInstaller,
                            ([System.Environment]::Is64BitProcess -and ($_.PSPath -notmatch '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node'))
                        )

                        # Build out an object and return it to the pipeline if there's no filterscript or the filterscript returns something.
                        if (!$FilterScript -or (& $Script:CommandTable.'ForEach-Object' -InputObject $app -Process $FilterScript -ErrorAction Ignore))
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Found installed application [$($app.DisplayName)]$(if ($app.DisplayVersion) {" version [$($app.DisplayVersion)]"})."
                            return $app
                        }
                    }
                }

                # Write to log the number of entries skipped due to them being considered updates.
                if (!$IncludeUpdatesAndHotfixes -and $updatesSkippedCounter)
                {
                    if ($updatesSkippedCounter -eq 1)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Skipped 1 entry while searching, because it was considered a Microsoft update.'
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Skipped $UpdatesSkippedCounter entries while searching, because they were considered Microsoft updates."
                    }
                }

                # Return any accumulated apps to the caller.
                if ($installedApplication)
                {
                    return $installedApplication
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Found no application based on the supplied FilterScript.'
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#---------------------------------------------------------------------------
#
# MARK: Get-ADTBoundParametersAndDefaultValues
#
#---------------------------------------------------------------------------

function Get-ADTBoundParametersAndDefaultValues
{
    <#
    .SYNOPSIS
        Returns a hashtable with the output of $PSBoundParameters and default-valued parameters for the given InvocationInfo.

    .DESCRIPTION
        This function processes the provided InvocationInfo and combines the results of $PSBoundParameters and default-valued parameters via the InvocationInfo's ScriptBlock AST (Abstract Syntax Tree).

    .PARAMETER Invocation
        The script or function's InvocationInfo ($MyInvocation) to process.

    .PARAMETER ParameterSetName
        The ParameterSetName to use as a filter against the Invocation's parameters.

    .PARAMETER HelpMessage
        The HelpMessage field to use as a filter against the Invocation's parameters.

    .PARAMETER Exclude
        One or more parameter names to exclude from the results.

    .PARAMETER CommonParameters
        Specifies whether PowerShell advanced function common parameters should be included.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.Generic.Dictionary[System.String, System.Object]

        Get-ADTBoundParametersAndDefaultValues returns a dictionary of the same base type as $PSBoundParameters for API consistency.

    .EXAMPLE
        Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation

        Returns a $PSBoundParameters-compatible dictionary with the bound parameters and any default values.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ParameterSetName', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'HelpMessage', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Exclude', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[System.String, System.Object]])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.InvocationInfo]$Invocation,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ParameterSetName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$HelpMessage,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Exclude,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CommonParameters
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Hold array of common parameters for filtration.
        $commonParams = if (!$CommonParameters)
        {
            $(
                [System.Management.Automation.PSCmdlet]::CommonParameters
                [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            )
        }

        # Internal function for testing parameter attributes.
        function Test-NamedAttributeArgumentAst
        {
            [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Argument', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
            [CmdletBinding()]
            [OutputType([System.Boolean])]
            param
            (
                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.Management.Automation.Language.ParameterAst]$Parameter,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.String]$Argument,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.String]$Value
            )

            # Test whether we have AttributeAst objects.
            if (!($attributes = $Parameter.Attributes | & { process { if ($_ -is [System.Management.Automation.Language.AttributeAst]) { return $_ } } }))
            {
                return $false
            }

            # Test whether we have NamedAttributeArgumentAst objects.
            if (!($namedArguments = $attributes.NamedArguments | & { process { if ($_.ArgumentName.Equals($Argument)) { return $_ } } }))
            {
                return $false
            }

            # Test whether any NamedAttributeArgumentAst objects match our value.
            return $namedArguments.Argument.Value.Contains($Value)
        }
    }

    process
    {
        try
        {
            try
            {
                # Get the parameters from the provided invocation. This can vary between simple/advanced functions and scripts.
                $parameters = if ($Invocation.MyCommand.ScriptBlock.Ast -is [System.Management.Automation.Language.FunctionDefinitionAst])
                {
                    # Test whether this is a simple or advanced function.
                    if ($Invocation.MyCommand.ScriptBlock.Ast.Parameters -and $Invocation.MyCommand.ScriptBlock.Ast.Parameters.Count)
                    {
                        $Invocation.MyCommand.ScriptBlock.Ast.Parameters
                    }
                    elseif ($Invocation.MyCommand.ScriptBlock.Ast.Body.ParamBlock -and $Invocation.MyCommand.ScriptBlock.Ast.Body.ParamBlock.Parameters.Count)
                    {
                        $Invocation.MyCommand.ScriptBlock.Ast.Body.ParamBlock.Parameters
                    }
                }
                elseif ($Invocation.MyCommand.ScriptBlock.Ast.ParamBlock -and $Invocation.MyCommand.ScriptBlock.Ast.ParamBlock.Parameters.Count)
                {
                    $Invocation.MyCommand.ScriptBlock.Ast.ParamBlock.Parameters
                }

                # Throw if we don't have any parameters at all.
                if (!$parameters -or !$parameters.Count)
                {
                    $naerParams = @{
                        Exception = [System.InvalidOperationException]::new("Unable to find parameters within the provided invocation's scriptblock AST.")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                        ErrorId = 'InvocationParametersNotFound'
                        TargetObject = $Invocation.MyCommand.ScriptBlock.Ast
                        RecommendedAction = "Please verify your function or script parameter configuration and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Open dictionary to store all params and their values to return.
                $obj = [System.Collections.Generic.Dictionary[System.String, System.Object]]::new()

                # Inject our already bound parameters into above object.
                $Invocation.BoundParameters.GetEnumerator() | & {
                    process
                    {
                        # Filter out common parameters.
                        if ($commonParams -notcontains $_.Key)
                        {
                            $obj.Add($_.Key, $_.Value)
                        }
                    }
                }

                # Build out the dictionary for returning.
                $parameters | & {
                    process
                    {
                        # Filter out excluded values.
                        if ($Exclude -contains $_.Name.VariablePath.UserPath)
                        {
                            $null = $obj.Remove($_.Name.VariablePath.UserPath)
                            return
                        }

                        # Filter out values based on the specified parameter set.
                        if ($ParameterSetName -and !(Test-NamedAttributeArgumentAst -Parameter $_ -Argument ParameterSetName -Value $ParameterSetName))
                        {
                            $null = $obj.Remove($_.Name.VariablePath.UserPath)
                            return
                        }

                        # Filter out values based on the specified help message.
                        if ($HelpMessage -and !(Test-NamedAttributeArgumentAst -Parameter $_ -Argument HelpMessage -Value $HelpMessage))
                        {
                            $null = $obj.Remove($_.Name.VariablePath.UserPath)
                            return
                        }

                        # Filter out parameters already bound.
                        if ($obj.ContainsKey($_.Name.VariablePath.UserPath))
                        {
                            return
                        }

                        # Filter out parameters without a default value.
                        if ($null -eq $_.DefaultValue)
                        {
                            return
                        }

                        # Add the parameter and its value.
                        $obj.Add($_.Name.VariablePath.UserPath, $_.DefaultValue.SafeGetValue())
                    }
                }

                # Return dictionary to the caller, even if it's empty.
                return $obj
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTCommandTable
#
#-----------------------------------------------------------------------------

function Get-ADTCommandTable
{
    <#
    .SYNOPSIS
        Returns PSAppDeployToolkit's safe command lookup table.

    .DESCRIPTION
        This function returns PSAppDeployToolkit's safe command lookup table, which can be used for command lookups within extending modules.

        Please note that PSAppDeployToolkit's safe command table only has commands in it that are used within this module, and not necessarily all commands offered by PowerShell and its built-in modules out of the box.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.ObjectModel.ReadOnlyDictionary[System.String, System.Management.Automation.CommandInfo]

        Returns PSAppDeployTookit's safe command lookup table as a ReadOnlyDictionary.

    .EXAMPLE
        Get-ADTCommandTable

        Returns PSAppDeployToolkit's safe command lookup table.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    # Return the module's read-only CommandTable to the caller.
    return $Script:CommandTable
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTConfig
#
#-----------------------------------------------------------------------------

function Get-ADTConfig
{
    <#
    .SYNOPSIS
        Retrieves the configuration data for the ADT module.

    .DESCRIPTION
        The Get-ADTConfig function retrieves the configuration data for the ADT module. This function ensures that the ADT module has been initialized before attempting to retrieve the configuration data. If the module is not initialized, it throws an error.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.Hashtable

        Returns the configuration data as a hashtable.

    .EXAMPLE
        $config = Get-ADTConfig

        This example retrieves the configuration data for the ADT module and stores it in the $config variable.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    # Return the config database if initialized.
    if (!$Script:ADT.Config -or !$Script:ADT.Config.Count)
    {
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("Please ensure that [Initialize-ADTModule] is called before using any $($MyInvocation.MyCommand.Module.Name) functions.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
            ErrorId = 'ADTConfigNotLoaded'
            TargetObject = $Script:ADT.Config
            RecommendedAction = "Please ensure the module is initialized via [Initialize-ADTModule] and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }
    return $Script:ADT.Config
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTDeferHistory
#
#-----------------------------------------------------------------------------

function Get-ADTDeferHistory
{
    <#
    .SYNOPSIS
        Get the history of deferrals in the registry for the current application.

    .DESCRIPTION
        Get the history of deferrals in the registry for the current application.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Get-DeferHistory

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com

    #>

    [CmdletBinding()]
    param
    (
    )

    try
    {
        (& $Script:CommandTable.'Get-ADTSession').GetDeferHistory()
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTEnvironment
#
#-----------------------------------------------------------------------------

function Get-ADTEnvironment
{
    <#
    .SYNOPSIS
        Retrieves the environment data for the ADT module. This function has been replaced by Get-ADTEnvironmentTable and will be removed from a future release.

    .DESCRIPTION
        The Get-ADTEnvironment function retrieves the environment data for the ADT module. This function ensures that the ADT module has been initialized before attempting to retrieve the environment data. If the module is not initialized, it throws an error.

        This function has been replaced by Get-ADTEnvironmentTable and will be removed from a future release.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.Specialized.OrderedDictionary

        Returns the environment data as a read-only ordered dictionary.

    .EXAMPLE
        $environment = Get-ADTEnvironment

        This example retrieves the environment data for the ADT module and stores it in the $environment variable.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    # Announce deprecation and return the environment database if initialized.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The function [$($MyInvocation.MyCommand.Name)] has been replaced by [Get-ADTEnvironmentTable]. Please migrate your scripts as this will be removed in a future update." -Severity 2
    return (& $Script:CommandTable.'Get-ADTEnvironmentTable')
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTEnvironmentTable
#
#-----------------------------------------------------------------------------

function Get-ADTEnvironmentTable
{
    <#
    .SYNOPSIS
        Retrieves the environment data for the ADT module.

    .DESCRIPTION
        The Get-ADTEnvironmentTable function retrieves the environment data for the ADT module. This function ensures that the ADT module has been initialized before attempting to retrieve the environment data. If the module is not initialized, it throws an error.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.Specialized.OrderedDictionary

        Returns the environment data as a read-only ordered dictionary.

    .EXAMPLE
        $environment = Get-ADTEnvironmentTable

        This example retrieves the environment data for the ADT module and stores it in the $environment variable.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    # Return the environment database if initialized.
    if (!$Script:ADT.Environment -or !$Script:ADT.Environment.Count)
    {
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("Please ensure that [Initialize-ADTModule] is called before using any $($MyInvocation.MyCommand.Module.Name) functions.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
            ErrorId = 'ADTEnvironmentDatabaseEmpty'
            TargetObject = $Script:ADT.Environment
            RecommendedAction = "Please ensure the module is initialized via [Initialize-ADTModule] and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }
    return $Script:ADT.Environment
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTFileVersion
#
#-----------------------------------------------------------------------------

function Get-ADTFileVersion
{
    <#
    .SYNOPSIS
        Gets the version of the specified file.

    .DESCRIPTION
        The Get-ADTFileVersion function retrieves the version information of the specified file. By default, it returns the FileVersion, but it can also return the ProductVersion if the -ProductVersion switch is specified.

    .PARAMETER File
        The path of the file.

    .PARAMETER ProductVersion
        Switch that makes the command return ProductVersion instead of FileVersion.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the version of the specified file.

    .EXAMPLE
        Get-ADTFileVersion -File "$env:ProgramFilesX86\Adobe\Reader 11.0\Reader\AcroRd32.exe"

        This example retrieves the FileVersion of the specified Adobe Reader executable.

    .EXAMPLE
        Get-ADTFileVersion -File "$env:ProgramFilesX86\Adobe\Reader 11.0\Reader\AcroRd32.exe" -ProductVersion

        This example retrieves the ProductVersion of the specified Adobe Reader executable.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!$_.Exists)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName File -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
                }
                if (!$_.VersionInfo -or (!$_.VersionInfo.FileVersion -and !$_.VersionInfo.ProductVersion))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName File -ProvidedValue $_ -ExceptionMessage 'The specified file does not have any version info.'))
                }
                return !!$_.VersionInfo
            })]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ProductVersion
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        if ($ProductVersion)
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Product version is [$($File.VersionInfo.ProductVersion)]."
            return $File.VersionInfo.ProductVersion.Trim()
        }
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "File version is [$($File.VersionInfo.FileVersion)]."
        return $File.VersionInfo.FileVersion.Trim()
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTFreeDiskSpace
#
#-----------------------------------------------------------------------------

function Get-ADTFreeDiskSpace
{
    <#
    .SYNOPSIS
        Retrieves the free disk space in MB on a particular drive (defaults to system drive).

    .DESCRIPTION
        The Get-ADTFreeDiskSpace function retrieves the free disk space in MB on a specified drive. If no drive is specified, it defaults to the system drive. This function is useful for monitoring disk space availability.

    .PARAMETER Drive
        The drive to check free disk space on.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Double

        Returns the free disk space in MB.

    .EXAMPLE
        Get-ADTFreeDiskSpace -Drive 'C:'

        This example retrieves the free disk space on the C: drive.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if (!$_.TotalSize)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Drive -ProvidedValue $_ -ExceptionMessage 'The specified drive does not exist or has no media loaded.'))
                }
                return !!$_.TotalSize
            })]
        [System.IO.DriveInfo]$Drive = [System.IO.Path]::GetPathRoot([System.Environment]::SystemDirectory)
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Retrieving free disk space for drive [$Drive]."
        $freeDiskSpace = [System.Math]::Round($Drive.AvailableFreeSpace / 1MB)
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Free disk space for drive [$Drive]: [$freeDiskSpace MB]."
        return $freeDiskSpace
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTIniValue
#
#-----------------------------------------------------------------------------

function Get-ADTIniValue
{
    <#
    .SYNOPSIS
        Parses an INI file and returns the value of the specified section and key.

    .DESCRIPTION
        The Get-ADTIniValue function parses an INI file and returns the value of the specified section and key. This function is useful for retrieving configuration settings stored in INI files.

    .PARAMETER FilePath
        Path to the INI file.

    .PARAMETER Section
        Section within the INI file.

    .PARAMETER Key
        Key within the section of the INI file.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the value of the specified section and key.

    .EXAMPLE
        Get-ADTIniValue -FilePath "$env:ProgramFilesX86\IBM\Notes\notes.ini" -Section 'Notes' -Key 'KeyFileName'

        This example retrieves the value of the 'KeyFileName' key in the 'Notes' section of the specified INI file.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Section,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Reading INI Key: [Section = $Section] [Key = $Key]."
        try
        {
            try
            {
                $iniValue = [PSADT.Configuration.IniFile]::GetSectionKeyValue($Section, $Key, $FilePath)
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "INI Key Value: [Section = $Section] [Key = $Key] [Value = $iniValue]."
                return $iniValue
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to read INI file key value."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTLoggedOnUser
#
#-----------------------------------------------------------------------------

function Get-ADTLoggedOnUser
{
    <#
    .SYNOPSIS
        Retrieves session details for all local and RDP logged on users.

    .DESCRIPTION
        The Get-ADTLoggedOnUser function retrieves session details for all local and RDP logged on users using Win32 APIs. It provides information such as NTAccount, SID, UserName, DomainName, SessionId, SessionName, ConnectState, IsCurrentSession, IsConsoleSession, IsUserSession, IsActiveUserSession, IsRdpSession, IsLocalAdmin, LogonTime, IdleTime, DisconnectTime, ClientName, ClientProtocolType, ClientDirectory, and ClientBuildNumber.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.UserSessionInfo

        Returns a custom type with information about user sessions:
        - NTAccount
        - SID
        - UserName
        - DomainName
        - SessionId
        - SessionName
        - ConnectState
        - IsCurrentSession
        - IsConsoleSession
        - IsUserSession
        - IsActiveUserSession
        - IsRdpSession
        - IsLocalAdmin
        - LogonTime
        - IdleTime
        - DisconnectTime
        - ClientName
        - ClientProtocolType
        - ClientDirectory
        - ClientBuildNumber

    .EXAMPLE
        Get-ADTLoggedOnUser

        This example retrieves session details for all local and RDP logged on users.

    .NOTES
        An active ADT session is NOT required to use this function.

        Description of ConnectState property:

        Value        Description
        -----        -----------
        Active       A user is logged on to the session.
        ConnectQuery The session is in the process of connecting to a client.
        Connected    A client is connected to the session.
        Disconnected The session is active, but the client has disconnected from it.
        Down         The session is down due to an error.
        Idle         The session is waiting for a client to connect.
        Initializing The session is initializing.
        Listening    The session is listening for connections.
        Reset        The session is being reset.
        Shadowing    This session is shadowing another session.

        Description of IsActiveUserSession property:
        - If a console user exists, then that will be the active user session.
        - If no console user exists but users are logged in, such as on terminal servers, then the first logged-in non-console user that has ConnectState either 'Active' or 'Connected' is the active user.

        Description of IsRdpSession property:
        - Gets a value indicating whether the user is associated with an RDP client session.

        Description of IsLocalAdmin property:
        - Checks whether the user is a member of the Administrators group

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Collections.ObjectModel.ReadOnlyCollection[PSADT.WTSSession.CompatibilitySessionInfo]])]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Getting session information for all logged on users.'
        try
        {
            try
            {
                return [PSADT.WTSSession.SessionManager]::GetCompatibilitySessionInfo()
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTMsiExitCodeMessage
#
#-----------------------------------------------------------------------------

function Get-ADTMsiExitCodeMessage
{
    <#
    .SYNOPSIS
        Get message for MSI exit code.

    .DESCRIPTION
        Get message for MSI exit code by reading it from msimsg.dll.

    .PARAMETER MsiExitCode
        MSI exit code.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the message for the MSI exit code.

    .EXAMPLE
        Get-ADTMsiExitCodeMessage -MsiExitCode 1618

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        http://msdn.microsoft.com/en-us/library/aa368542(v=vs.85).aspx

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$MsiExitCode
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Only return the output if we receive something from the library.
                if (![System.String]::IsNullOrWhiteSpace(($msg = [PSADT.Installer.Msi]::GetMessageFromMsiExitCode($MsiExitCode))))
                {
                    return $msg
                }
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTMsiTableProperty
#
#-----------------------------------------------------------------------------

function Get-ADTMsiTableProperty
{
    <#
    .SYNOPSIS
        Get all of the properties from a Windows Installer database table or the Summary Information stream and return as a custom object.

    .DESCRIPTION
        Use the Windows Installer object to read all of the properties from a Windows Installer database table or the Summary Information stream.

    .PARAMETER Path
        The fully qualified path to an database file. Supports .msi and .msp files.

    .PARAMETER TransformPath
        The fully qualified path to a list of MST file(s) which should be applied to the MSI file.

    .PARAMETER Table
        The name of the the MSI table from which all of the properties must be retrieved. Default is: 'Property'.

    .PARAMETER TablePropertyNameColumnNum
        Specify the table column number which contains the name of the properties. Default is: 1 for MSIs and 2 for MSPs.

    .PARAMETER TablePropertyValueColumnNum
        Specify the table column number which contains the value of the properties. Default is: 2 for MSIs and 3 for MSPs.

    .PARAMETER GetSummaryInformation
        Retrieves the Summary Information for the Windows Installer database.

        Summary Information property descriptions: https://msdn.microsoft.com/en-us/library/aa372049(v=vs.85).aspx

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Management.Automation.PSObject

        Returns a custom object with the following properties: 'Name' and 'Value'.

    .EXAMPLE
        Get-ADTMsiTableProperty -Path 'C:\Package\AppDeploy.msi' -TransformPath 'C:\Package\AppDeploy.mst'

        Retrieve all of the properties from the default 'Property' table.

    .EXAMPLE
        Get-ADTMsiTableProperty -Path 'C:\Package\AppDeploy.msi' -TransformPath 'C:\Package\AppDeploy.mst' -Table 'Property' | Select-Object -ExpandProperty ProductCode

        Retrieve all of the properties from the 'Property' table and then pipe to Select-Object to select the ProductCode property.

    .EXAMPLE
        Get-ADTMsiTableProperty -Path 'C:\Package\AppDeploy.msi' -GetSummaryInformation

        Retrieve the Summary Information for the Windows Installer database.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(DefaultParameterSetName = 'TableInfo')]
    [OutputType([System.Collections.ObjectModel.ReadOnlyDictionary[System.String, System.Object]])]
    [OutputType([PSADT.Types.MsiSummaryInfo])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!(& $Script:CommandTable.'Test-Path' -Path $_ -PathType Leaf))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if (!(& $Script:CommandTable.'Test-Path' -Path $_ -PathType Leaf))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName TransformPath -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String[]]$TransformPath,

        [Parameter(Mandatory = $false, ParameterSetName = 'TableInfo')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Table,

        [Parameter(Mandatory = $false, ParameterSetName = 'TableInfo')]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$TablePropertyNameColumnNum,

        [Parameter(Mandatory = $false, ParameterSetName = 'TableInfo')]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$TablePropertyValueColumnNum,

        [Parameter(Mandatory = $true, ParameterSetName = 'SummaryInfo')]
        [System.Management.Automation.SwitchParameter]$GetSummaryInformation
    )

    begin
    {
        # Set default values.
        if (!$PSBoundParameters.ContainsKey('Table'))
        {
            $Table = ('MsiPatchMetadata', 'Property')[[System.IO.Path]::GetExtension($Path) -eq '.msi']
        }
        if (!$PSBoundParameters.ContainsKey('TablePropertyNameColumnNum'))
        {
            $TablePropertyNameColumnNum = 2 - ([System.IO.Path]::GetExtension($Path) -eq '.msi')
        }
        if (!$PSBoundParameters.ContainsKey('TablePropertyValueColumnNum'))
        {
            $TablePropertyValueColumnNum = 3 - ([System.IO.Path]::GetExtension($Path) -eq '.msi')
        }

        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'TableInfo')
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Reading data from Windows Installer database file [$Path] in table [$Table]."
        }
        else
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Reading the Summary Information from the Windows Installer database file [$Path]."
        }
        try
        {
            try
            {
                # Create a Windows Installer object and define properties for how the MSI database is opened
                $Installer = & $Script:CommandTable.'New-Object' -ComObject WindowsInstaller.Installer
                $msiOpenDatabaseModeReadOnly = 0
                $msiSuppressApplyTransformErrors = 63
                $msiOpenDatabaseModePatchFile = 32
                $msiOpenDatabaseMode = if (($IsMspFile = [IO.Path]::GetExtension($Path) -eq '.msp'))
                {
                    $msiOpenDatabaseModePatchFile
                }
                else
                {
                    $msiOpenDatabaseModeReadOnly
                }

                # Open database in read only mode and apply a list of transform(s).
                $Database = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Installer -MethodName OpenDatabase -ArgumentList @($Path, $msiOpenDatabaseMode)
                if ($TransformPath -and !$IsMspFile)
                {
                    $null = foreach ($Transform in $TransformPath)
                    {
                        & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Database -MethodName ApplyTransform -ArgumentList @($Transform, $msiSuppressApplyTransformErrors)
                    }
                }

                # Get either the requested windows database table information or summary information.
                if ($GetSummaryInformation)
                {
                    # Get the SummaryInformation from the windows installer database.
                    # Summary property descriptions: https://msdn.microsoft.com/en-us/library/aa372049(v=vs.85).aspx
                    $SummaryInformation = & $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $Database -PropertyName SummaryInformation
                    return [PSADT.Types.MsiSummaryInfo]::new(
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(1)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(2)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(3)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(4)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(5)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(6)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(7)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(8)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(9)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(11)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(12)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(13)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(14)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(15)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(16)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(18)),
                        (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(19))
                    )
                }

                # Open the requested table view from the database.
                $TableProperties = [System.Collections.Generic.Dictionary[System.String, System.Object]]::new()
                $View = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Database -MethodName OpenView -ArgumentList @("SELECT * FROM $Table")
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Execute

                # Retrieve the first row from the requested table. If the first row was successfully retrieved, then save data and loop through the entire table.
                # https://msdn.microsoft.com/en-us/library/windows/desktop/aa371136(v=vs.85).aspx
                while (($Record = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Fetch))
                {
                    $TableProperties.Add((& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $Record -PropertyName StringData -ArgumentList @($TablePropertyNameColumnNum)), (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $Record -PropertyName StringData -ArgumentList @($TablePropertyValueColumnNum)))
                }

                # Return the accumulated results. We can't use a custom object for this as we have no idea what's going to be in the properties of a given MSI.
                # We also can't use a pscustomobject accelerator here as the MSI may have the same keys with different casing, necessitating the use of a dictionary for storage.
                if ($TableProperties.Count)
                {
                    return [System.Collections.ObjectModel.ReadOnlyDictionary[System.String, System.Object]]$TableProperties
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to get the MSI table [$Table]."
        }
        finally
        {
            # Release all COM objects to prevent file locks.
            $null = foreach ($variable in (& $Script:CommandTable.'Get-Variable' -Name View, SummaryInformation, Database, Installer -ValueOnly -ErrorAction Ignore))
            {
                try
                {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($variable)
                }
                catch
                {
                    $null
                }
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTObjectProperty
#
#-----------------------------------------------------------------------------

function Get-ADTObjectProperty
{
    <#
    .SYNOPSIS
        Get a property from any object.

    .DESCRIPTION
        Get a property from any object.

    .PARAMETER InputObject
        Specifies an object which has properties that can be retrieved.

    .PARAMETER PropertyName
        Specifies the name of a property to retrieve.

    .PARAMETER ArgumentList
        Argument to pass to the property being retrieved.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Object

        Returns the value of the property being retrieved.

    .EXAMPLE
        Get-ADTObjectProperty -InputObject $Record -PropertyName 'StringData' -ArgumentList @(1)

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

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

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                return $InputObject.GetType().InvokeMember($PropertyName, [Reflection.BindingFlags]::GetProperty, $null, $InputObject, $ArgumentList, $null, $null, $null)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTOperatingSystemInfo
#
#-----------------------------------------------------------------------------

function Get-ADTOperatingSystemInfo
{
    <#
    .SYNOPSIS
        Gets information about the current computer's operating system.

    .DESCRIPTION
        Gets information about the current computer's operating system, such as name, version, edition, and other information.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.OperatingSystem.OSVersionInfo

        Returns an PSADT.OperatingSystem.OSVersionInfo object containing the current computer's operating system information.

    .EXAMPLE
        Get-ADTOperatingSystemInfo

        Gets an PSADT.OperatingSystem.OSVersionInfo object containing the current computer's operating system information.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    return [PSADT.OperatingSystem.OSVersionInfo]::Current
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTPEFileArchitecture
#
#-----------------------------------------------------------------------------

function Get-ADTPEFileArchitecture
{
    <#
    .SYNOPSIS
        Determine if a PE file is a 32-bit or a 64-bit file.

    .DESCRIPTION
        Determine if a PE file is a 32-bit or a 64-bit file by examining the file's image file header.

        PE file extensions: .exe, .dll, .ocx, .drv, .sys, .scr, .efi, .cpl, .fon

    .PARAMETER FilePath
        Path to the PE file to examine.

    .PARAMETER PassThru
        Get the file object, attach a property indicating the file binary type, and write to pipeline.

    .INPUTS
        System.IO.FileInfo

        Accepts a FileInfo object from the pipeline.

    .OUTPUTS
        System.String

        Returns a string indicating the file binary type.

    .EXAMPLE
        Get-ADTPEFileArchitecture -FilePath "$env:windir\notepad.exe"

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    [OutputType([PSADT.Shared.SystemArchitecture])]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({
                if (!$_.Exists -or ($_ -notmatch '\.(exe|dll|ocx|drv|sys|scr|efi|cpl|fon)$'))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'One or more files either does not exist or has an invalid extension.'))
                }
                return !!$_
            })]
        [System.IO.FileInfo[]]$FilePath,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        [System.Int32]$MACHINE_OFFSET = 4
        [System.Int32]$PE_POINTER_OFFSET = 60
        [System.Byte[]]$data = [System.Byte[]]::new(4096)
    }

    process
    {
        foreach ($Path in $filePath)
        {
            try
            {
                try
                {
                    # Read the first 4096 bytes of the file.
                    $stream = [System.IO.FileStream]::new($Path.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                    $null = $stream.Read($data, 0, $data.Count)
                    $stream.Flush()
                    $stream.Close()

                    # Get the file header from the header's address, factoring in any offsets.
                    $peArchValue = [System.BitConverter]::ToUInt16($data, [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET) + $MACHINE_OFFSET)
                    $peArchEnum = [PSADT.Shared.SystemArchitecture]::Unknown; $null = [PSADT.Shared.SystemArchitecture]::TryParse($peArchValue, [ref]$peArchEnum)
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "File [$($Path.FullName)] has a detected file architecture of [$peArchEnum]."
                    if ($PassThru)
                    {
                        return ($Path | & $Script:CommandTable.'Add-Member' -MemberType NoteProperty -Name BinaryType -Value $peArchEnum -Force -PassThru)
                    }
                    return $peArchEnum
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTPendingReboot
#
#-----------------------------------------------------------------------------

function Get-ADTPendingReboot
{
    <#
    .SYNOPSIS
        Get the pending reboot status on a local computer.

    .DESCRIPTION
        Check WMI and the registry to determine if the system has a pending reboot operation from any of the following:
        a) Component Based Servicing (Vista, Windows 2008)
        b) Windows Update / Auto Update (XP, Windows 2003 / 2008)
        c) SCCM 2012 Clients (DetermineIfRebootPending WMI method)
        d) App-V Pending Tasks (global based Appv 5.0 SP2)
        e) Pending File Rename Operations (XP, Windows 2003 / 2008)

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.RebootInfo

        Returns a custom object with the following properties:
        - ComputerName
        - LastBootUpTime
        - IsSystemRebootPending
        - IsCBServicingRebootPending
        - IsWindowsUpdateRebootPending
        - IsSCCMClientRebootPending
        - IsIntuneClientRebootPending
        - IsFileRenameRebootPending
        - PendingFileRenameOperations
        - ErrorMsg

    .EXAMPLE
        Get-ADTPendingReboot

        This example retrieves the pending reboot status on the local computer and returns a custom object with detailed information.

    .EXAMPLE
        (Get-ADTPendingReboot).IsSystemRebootPending

        This example returns a boolean value determining whether or not there is a pending reboot operation.

    .NOTES
        An active ADT session is NOT required to use this function.

        ErrorMsg only contains something if an error occurred.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([PSADT.Types.RebootInfo])]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $PendRebootErrorMsg = [System.Collections.Specialized.StringCollection]::new()
        $HostName = [System.Net.Dns]::GetHostName()
    }

    process
    {
        try
        {
            try
            {
                # Get the date/time that the system last booted up.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Getting the pending reboot status on the local computer [$HostName]."
                $LastBootUpTime = [System.DateTime]::Now - [System.TimeSpan]::FromMilliseconds([System.Environment]::TickCount)

                # Determine if a Windows Vista/Server 2008 and above machine has a pending reboot from a Component Based Servicing (CBS) operation.
                $IsCBServicingRebootPending = & $Script:CommandTable.'Test-Path' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'

                # Determine if there is a pending reboot from a Windows Update.
                $IsWindowsUpdateRebootPending = & $Script:CommandTable.'Test-Path' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'

                # Determine if there is a pending reboot from an App-V global Pending Task. (User profile based tasks will complete on logoff/logon).
                $IsAppVRebootPending = & $Script:CommandTable.'Test-Path' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Software\Microsoft\AppV\Client\PendingTasks'

                # Get the value of PendingFileRenameOperations.
                $PendingFileRenameOperations = if ($IsFileRenameRebootPending = & $Script:CommandTable.'Test-ADTRegistryValue' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations')
                {
                    try
                    {
                        & $Script:CommandTable.'Get-ItemProperty' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager' | & $Script:CommandTable.'Select-Object' -ExpandProperty PendingFileRenameOperations
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to get PendingFileRenameOperations.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                        $null = $PendRebootErrorMsg.Add("Failed to get PendingFileRenameOperations: $($_.Exception.Message)")
                    }
                }

                # Determine SCCM 2012 Client reboot pending status.
                $IsSCCMClientRebootPending = try
                {
                    if (($SCCMClientRebootStatus = & $Script:CommandTable.'Invoke-CimMethod' -Namespace ROOT\CCM\ClientSDK -ClassName CCM_ClientUtilities -Name DetermineIfRebootPending).ReturnValue -eq 0)
                    {
                        $SCCMClientRebootStatus.IsHardRebootPending -or $SCCMClientRebootStatus.RebootPending
                    }
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to get IsSCCMClientRebootPending.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                    $null = $PendRebootErrorMsg.Add("Failed to get IsSCCMClientRebootPending: $($_.Exception.Message)")
                }

                # Determine Intune Management Extension reboot pending status.
                $IsIntuneClientRebootPending = try
                {
                    !!(& $Script:CommandTable.'Get-Item' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\RebootSettings\RebootFlag')
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to get IsIntuneClientRebootPending.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                    $null = $PendRebootErrorMsg.Add("Failed to get IsIntuneClientRebootPending: $($_.Exception.Message)")
                }

                # Create a custom object containing pending reboot information for the system.
                $PendingRebootInfo = [PSADT.Types.RebootInfo]::new(
                    $HostName,
                    $LastBootUpTime,
                    $IsCBServicingRebootPending -or $IsWindowsUpdateRebootPending -or $IsFileRenameRebootPending -or $IsSCCMClientRebootPending,
                    $IsCBServicingRebootPending,
                    $IsWindowsUpdateRebootPending,
                    $IsSCCMClientRebootPending,
                    $IsIntuneClientRebootPending,
                    $IsAppVRebootPending,
                    $IsFileRenameRebootPending,
                    $PendingFileRenameOperations,
                    $PendRebootErrorMsg
                )
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Pending reboot status on the local computer [$HostName]:`n$($PendingRebootInfo | & $Script:CommandTable.'Format-List' | & $Script:CommandTable.'Out-String' -Width ([System.Int32]::MaxValue))"
                return $PendingRebootInfo
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTPowerShellProcessPath
#
#-----------------------------------------------------------------------------

function Get-ADTPowerShellProcessPath
{
    <#
    .SYNOPSIS
        Retrieves the path to the PowerShell executable.

    .DESCRIPTION
        The Get-ADTPowerShellProcessPath function returns the path to the PowerShell executable. It determines whether the current PowerShell session is running in Windows PowerShell or PowerShell Core and returns the appropriate executable path.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the path to the PowerShell executable as a string.

    .EXAMPLE
        Get-ADTPowerShellProcessPath

        This example retrieves the path to the PowerShell executable for the current session.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    return "$PSHOME\$(('powershell.exe', 'pwsh.exe')[$PSVersionTable.PSEdition.Equals('Core')])"
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTPresentationSettingsEnabledUsers
#
#-----------------------------------------------------------------------------

function Get-ADTPresentationSettingsEnabledUsers
{
    <#
    .SYNOPSIS
        Tests whether any users have presentation mode enabled on their device.

    .DESCRIPTION
        Tests whether any users have presentation mode enabled on their device. This can be enabled via the PC's Mobility Settings, or with PresentationSettings.exe.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.UserProfile

        Returns one or more UserProfile objects of the users with presentation mode enabled on their device.

    .EXAMPLE
        Get-ADTPresentationSettingsEnabledUsers

        Checks whether any users users have presentation settings enabled on their device and returns an associated UserProfile object.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([PSADT.Types.UserProfile])]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Checking whether any logged on users are in presentation mode..."
        try
        {
            try
            {
                # Build out params for Invoke-ADTAllUsersRegistryAction.
                $iaauraParams = @{
                    ScriptBlock = { if (& $Script:CommandTable.'Get-ADTRegistryKey' -Key Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\MobilePC\AdaptableSettings\Activity -Name Activity -SID $_.SID) { return $_ } }
                    UserProfiles = & $Script:CommandTable.'Get-ADTUserProfiles' -ExcludeDefaultUser -InformationAction SilentlyContinue
                }

                # Return UserProfile objects for each user with "I am currently giving a presentation" enabled.
                if (($usersInPresentationMode = & $Script:CommandTable.'Invoke-ADTAllUsersRegistryAction' @iaauraParams -SkipUnloadedProfiles -InformationAction SilentlyContinue))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The following users are currently in presentation mode: ['$([System.String]::Join("', '", $usersInPresentationMode.NTAccount))']."
                    return $usersInPresentationMode
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "There are no logged on users in presentation mode."
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTRegistryKey
#
#-----------------------------------------------------------------------------

function Get-ADTRegistryKey
{
    <#
    .SYNOPSIS
        Retrieves value names and value data for a specified registry key or optionally, a specific value.

    .DESCRIPTION
        Retrieves value names and value data for a specified registry key or optionally, a specific value.
        If the registry key does not exist or contain any values, the function will return $null by default.
        To test for existence of a registry key path, use built-in Test-Path cmdlet.

    .PARAMETER Key
        Path of the registry key.

    .PARAMETER Name
        Value name to retrieve (optional).

    .PARAMETER Wow6432Node
        Specify this switch to read the 32-bit registry (Wow6432Node) on 64-bit systems.

    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.
        Specify this parameter from the Invoke-ADTAllUsersRegistryAction function to read/edit HKCU registry settings for all users on the system.

    .PARAMETER ReturnEmptyKeyIfExists
        Return the registry key if it exists but it has no property/value pairs underneath it. Default is: $false.

    .PARAMETER DoNotExpandEnvironmentNames
        Return unexpanded REG_EXPAND_SZ values. Default is: $false.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the value of the registry key or value.

    .EXAMPLE
        Get-ADTRegistryKey -Key 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1AD147D0-BE0E-3D6C-AC11-64F6DC4163F1}'

        This example retrieves all value names and data for the specified registry key.

    .EXAMPLE
        Get-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\iexplore.exe'

        This example retrieves all value names and data for the specified registry key.

    .EXAMPLE
        Get-ADTRegistryKey -Key 'HKLM:Software\Wow6432Node\Microsoft\Microsoft SQL Server Compact Edition\v3.5' -Name 'Version'

        This example retrieves the 'Version' value data for the specified registry key.

    .EXAMPLE
        Get-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Name 'Path' -DoNotExpandEnvironmentNames

        This example retrieves the 'Path' value data without expanding environment variables.

    .EXAMPLE
        Get-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Example' -Name '(Default)'

        This example retrieves the default value data for the specified registry key.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ReturnEmptyKeyIfExists,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DoNotExpandEnvironmentNames
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        try
        {
            try
            {
                # If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID.
                $Key = if ($PSBoundParameters.ContainsKey('SID'))
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -Wow6432Node:$Wow6432Node -SID $SID
                }
                else
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -Wow6432Node:$Wow6432Node
                }

                # Check if the registry key exists before continuing.
                if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $Key))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Registry key [$Key] does not exist. Return `$null." -Severity 2
                    return
                }

                if ($PSBoundParameters.ContainsKey('Name'))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Getting registry key [$Key] value [$Name]."
                }
                else
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Getting registry key [$Key] and all property values."
                }

                # Get all property values for registry key.
                $regKeyValue = & $Script:CommandTable.'Get-ItemProperty' -LiteralPath $Key
                $regKeyValuePropertyCount = $regKeyValue | & $Script:CommandTable.'Measure-Object' | & $Script:CommandTable.'Select-Object' -ExpandProperty Count

                # Select requested property.
                if ($PSBoundParameters.ContainsKey('Name'))
                {
                    # Get the Value (do not make a strongly typed variable because it depends entirely on what kind of value is being read)
                    if ((& $Script:CommandTable.'Get-Item' -LiteralPath $Key | & $Script:CommandTable.'Select-Object' -ExpandProperty Property -ErrorAction Ignore) -notcontains $Name)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Registry key value [$Key] [$Name] does not exist. Return `$null."
                        return
                    }
                    if ($DoNotExpandEnvironmentNames)
                    {
                        # Only useful on 'ExpandString' values.
                        if ($Name -like '(Default)')
                        {
                            return (& $Script:CommandTable.'Get-Item' -LiteralPath $Key).GetValue($null, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                        }
                        else
                        {
                            return (& $Script:CommandTable.'Get-Item' -LiteralPath $Key).GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                        }
                    }
                    elseif ($Name -like '(Default)')
                    {
                        return (& $Script:CommandTable.'Get-Item' -LiteralPath $Key).GetValue($null)
                    }
                    else
                    {
                        return $regKeyValue | & $Script:CommandTable.'Select-Object' -ExpandProperty $Name
                    }
                }
                elseif ($regKeyValuePropertyCount -eq 0)
                {
                    # Select all properties or return empty key object.
                    if ($ReturnEmptyKeyIfExists)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "No property values found for registry key. Return empty registry key object [$Key]."
                        return (& $Script:CommandTable.'Get-Item' -LiteralPath $Key -Force)
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "No property values found for registry key. Return `$null."
                        return
                    }
                }

                # Return the populated registry key to the caller.
                return $regKeyValue
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to read registry key [$Key]$(if ($Name) {" value [$Name]"})."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTRunAsActiveUser
#
#-----------------------------------------------------------------------------

function Get-ADTRunAsActiveUser
{
    <#
    .SYNOPSIS
        Retrieves the active user session information.

    .DESCRIPTION
        The Get-ADTRunAsActiveUser function determines the account that will be used to execute commands in the user session when the toolkit is running under the SYSTEM account.
        The active console user will be chosen first. If no active console user is found, for multi-session operating systems, the first logged-on user will be used instead.

    .PARAMETER UserSessionInfo
        An array of UserSessionInfo objects to enumerate through. If not supplied, a fresh query will be performed.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.UserSessionInfo

        Returns a custom object containing the user session information.

    .EXAMPLE
        Get-ADTRunAsActiveUser

        This example retrieves the active user session information.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [PSADT.WTSSession.CompatibilitySessionInfo[]]$UserSessionInfo = (& $Script:CommandTable.'Get-ADTLoggedOnUser')
    )

    # Determine the account that will be used to execute commands in the user session when toolkit is running under the SYSTEM account.
    # The active console user will be chosen first. Failing that, for multi-session operating systems, the first logged on user will be used instead.
    try
    {
        $sessionInfoMember = if (& $Script:CommandTable.'Test-ADTIsMultiSessionOS') { 'IsCurrentSession' } else { 'IsActiveUserSession' }
        foreach ($userSessionInfo in $UserSessionInfo)
        {
            if ($userSessionInfo.NTAccount -and $userSessionInfo.$sessionInfoMember)
            {
                return $userSessionInfo
            }
        }
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTSchedulerTask
#
#-----------------------------------------------------------------------------

function Get-ADTSchedulerTask
{
    <#
    .SYNOPSIS
        Retrieve all details for scheduled tasks on the local computer.

    .DESCRIPTION
        Retrieve all details for scheduled tasks on the local computer using schtasks.exe. All property names have spaces and colons removed.
        This function is deprecated. Please migrate your scripts to use the built-in Get-ScheduledTask Cmdlet.

    .PARAMETER TaskName
        Specify the name of the scheduled task to retrieve details for. Uses regex match to find scheduled task.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.PSObject

        This function returns a PSObject with all scheduled task properties.

    .EXAMPLE
        Get-ADTSchedulerTask

        This example retrieves a list of all scheduled task properties.

    .EXAMPLE
        Get-ADTSchedulerTask | Out-GridView

        This example displays a grid view of all scheduled task properties.

    .EXAMPLE
        Get-ADTSchedulerTask | Select-Object -Property TaskName

        This example displays a list of all scheduled task names.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'TaskName', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TaskName
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Advise that this function is considered deprecated.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "The function [$($MyInvocation.MyCommand.Name)] is deprecated. Please migrate your scripts to use the built-in [Get-ScheduledTask] Cmdlet." -Severity 2
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Retrieving Scheduled Tasks...'
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
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Convert CSV data to objects and re-process to remove non-word characters before returning data to the caller.
                if (($schTasks = $exeSchtasksResults | & $Script:CommandTable.'ConvertFrom-Csv' | & { process { if (($_.TaskName -match '^\\') -and ([string]::IsNullOrWhiteSpace($TaskName) -or $_.TaskName -match $TaskName)) { return $_ } } }))
                {
                    return $schTasks | & $Script:CommandTable.'Select-Object' -Property ($schTasks[0].PSObject.Properties.Name | & {
                            process
                            {
                                @{ Label = $_ -replace '[^\w]'; Expression = [scriptblock]::Create("`$_.'$_'") }
                            }
                        })
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to retrieve scheduled tasks."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTServiceStartMode
#
#-----------------------------------------------------------------------------

function Get-ADTServiceStartMode
{
    <#
    .SYNOPSIS
        Retrieves the startup mode of a specified service.

    .DESCRIPTION
        Retrieves the startup mode of a specified service. This function checks the service's start type and adjusts the result if the service is set to 'Automatic (Delayed Start)'.

    .PARAMETER Service
        Specify the service object to retrieve the startup mode for.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the startup mode of the specified service.

    .EXAMPLE
        Get-ADTServiceStartMode -Service (Get-Service -Name 'wuauserv')

        Retrieves the startup mode of the 'wuauserv' service.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!$_.Name)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Service -ProvidedValue $_ -ExceptionMessage 'The specified service does not exist.'))
                }
                return !!$_
            })]
        [System.ServiceProcess.ServiceController]$Service
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Getting the service [$($Service.Name)] startup mode."
        try
        {
            try
            {
                # Get the start mode and adjust it if the automatic type is delayed.
                if ((($serviceStartMode = $Service.StartType) -eq 'Automatic') -and ((& $Script:CommandTable.'Get-ItemProperty' -LiteralPath "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\$($Service.Name)" -ErrorAction Ignore | & $Script:CommandTable.'Select-Object' -ExpandProperty DelayedAutoStart -ErrorAction Ignore) -eq 1))
                {
                    $serviceStartMode = 'Automatic (Delayed Start)'
                }

                # Return startup type to the caller.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Service [$($Service.Name)] startup mode is set to [$serviceStartMode]."
                return $serviceStartMode
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTSession
#
#-----------------------------------------------------------------------------

function Get-ADTSession
{
    <#
    .SYNOPSIS
        Retrieves the most recent ADT session.

    .DESCRIPTION
        The Get-ADTSession function returns the most recent session from the ADT module data. If no sessions are found, it throws an error indicating that an ADT session should be opened using Open-ADTSession before calling this function.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        ADTSession

        Returns the most recent session object from the ADT module data.

    .EXAMPLE
        Get-ADTSession

        This example retrieves the most recent ADT session.

    .NOTES
        An active ADT session is required to use this function.

        Requires: PSADT session should be initialized using Open-ADTSession

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    # Return the most recent session in the database.
    if (!$Script:ADT.Sessions.Count)
    {
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("Please ensure that [Open-ADTSession] is called before using any $($MyInvocation.MyCommand.Module.Name) functions.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
            ErrorId = 'ADTSessionBufferEmpty'
            TargetObject = $Script:ADT.Sessions
            RecommendedAction = "Please ensure a session is opened via [Open-ADTSession] and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }
    return $Script:ADT.Sessions[-1]
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTShortcut
#
#-----------------------------------------------------------------------------

function Get-ADTShortcut
{
    <#
    .SYNOPSIS
        Get information from a .lnk or .url type shortcut.

    .DESCRIPTION
        Get information from a .lnk or .url type shortcut. Returns a hashtable with details about the shortcut such as TargetPath, Arguments, Description, and more.

    .PARAMETER Path
        Path to the shortcut to get information from.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Collections.Hashtable

        Returns a hashtable with the following keys:
        - TargetPath
        - Arguments
        - Description
        - WorkingDirectory
        - WindowStyle
        - Hotkey
        - IconLocation
        - IconIndex
        - RunAsAdmin

    .EXAMPLE
        Get-ADTShortcut -Path "$envProgramData\Microsoft\Windows\Start Menu\My Shortcut.lnk"

        Retrieves information from the specified .lnk shortcut.

    .NOTES
        An active ADT session is NOT required to use this function.

        Url shortcuts only support TargetPath, IconLocation, and IconIndex.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([PSADT.Types.ShortcutUrl])]
    [OutputType([PSADT.Types.ShortcutLnk])]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_) -or (![System.IO.Path]::GetExtension($Path).ToLower().Equals('.lnk') -and ![System.IO.Path]::GetExtension($Path).ToLower().Equals('.url')))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist or does not have the correct extension.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$Path
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        # Make sure .NET's current directory is synced with PowerShell's.
        try
        {
            try
            {
                [System.IO.Directory]::SetCurrentDirectory((& $Script:CommandTable.'Get-Location' -PSProvider FileSystem).ProviderPath)
                $Output = @{ Path = [System.IO.Path]::GetFullPath($Path); TargetPath = $null; IconIndex = $null; IconLocation = $null }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Specified path [$Path] is not valid."
            return
        }

        try
        {
            try
            {
                # Build out remainder of object.
                if ($Path -match '\.url$')
                {
                    [System.IO.File]::ReadAllLines($Path) | & {
                        process
                        {
                            switch ($_)
                            {
                                { $_.StartsWith('URL=') } { $Output.TargetPath = $_.Replace('URL=', $null); break }
                                { $_.StartsWith('IconIndex=') } { $Output.IconIndex = $_.Replace('IconIndex=', $null); break }
                                { $_.StartsWith('IconFile=') } { $Output.IconLocation = $_.Replace('IconFile=', $null); break }
                            }
                        }
                    }
                    return [PSADT.Types.ShortcutUrl]::new(
                        $Output.Path,
                        $Output.TargetPath,
                        $Output.IconIndex,
                        $Output.IconLocation
                    )
                }
                else
                {
                    $shortcut = [System.Activator]::CreateInstance([System.Type]::GetTypeFromProgID('WScript.Shell')).CreateShortcut($FullPath)
                    $Output.IconLocation, $Output.IconIndex = $shortcut.IconLocation.Split(',')
                    return [PSADT.Types.ShortcutLnk]::new(
                        $Output.Path,
                        $shortcut.TargetPath,
                        $Output.IconIndex,
                        $Output.IconLocation,
                        $shortcut.Arguments,
                        $shortcut.Description,
                        $shortcut.WorkingDirectory,
                        $(switch ($shortcut.WindowStyle)
                            {
                                1 { 'Normal'; break }
                                3 { 'Maximized'; break }
                                7 { 'Minimized'; break }
                                default { 'Normal'; break }
                            }),
                        $shortcut.Hotkey,
                        !!([Systen.IO.FIle]::ReadAllBytes($FullPath)[21] -band 32)
                    )
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to read the shortcut [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTStringTable
#
#-----------------------------------------------------------------------------

function Get-ADTStringTable
{
    <#
    .SYNOPSIS
        Retrieves the string database from the ADT module.

    .DESCRIPTION
        The Get-ADTStringTable function returns the string database if it has been initialized. If the string database is not initialized, it throws an error indicating that Initialize-ADTModule should be called before using this function.

    .INPUTS
        None

        This function does not take any pipeline input.

    .OUTPUTS
        System.Collections.Hashtable

        Returns a hashtable containing the string database.

    .EXAMPLE
        Get-ADTStringTable

        This example retrieves the string database from the ADT module.

    .NOTES
        An active ADT session is NOT required to use this function.

        Requires: The module should be initialized using Initialize-ADTModule

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    # Return the string database if initialized.
    if (!$Script:ADT.Strings -or !$Script:ADT.Strings.Count)
    {
        $naerParams = @{
            Exception = [System.InvalidOperationException]::new("Please ensure that [Initialize-ADTModule] is called before using any $($MyInvocation.MyCommand.Module.Name) functions.")
            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
            ErrorId = 'ADTStringTableNotInitialized'
            TargetObject = $Script:ADT.Strings
            RecommendedAction = "Please ensure the module is initialized via [Initialize-ADTModule] and try again."
        }
        $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
    }
    return $Script:ADT.Strings
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTUniversalDate
#
#-----------------------------------------------------------------------------

function Get-ADTUniversalDate
{
    <#
    .SYNOPSIS
        Returns the date/time for the local culture in a universal sortable date time pattern.

    .DESCRIPTION
        Converts the current datetime or a datetime string for the current culture into a universal sortable date time pattern, e.g. 2013-08-22 11:51:52Z.

    .PARAMETER DateTime
        Specify the DateTime in the current culture.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the date/time for the local culture in a universal sortable date time pattern.

    .EXAMPLE
        Get-ADTUniversalDate

        Returns the current date in a universal sortable date time pattern.

    .EXAMPLE
        Get-ADTUniversalDate -DateTime '25/08/2013'

        Returns the date for the current culture in a universal sortable date time pattern.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DateTime = [System.DateTime]::Now.ToString([System.Globalization.DateTimeFormatInfo]::CurrentInfo.UniversalSortableDateTimePattern)
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Remove any tailing Z, otherwise it could get converted to a different time zone. Then, convert the date to a universal sortable date time pattern based on the current culture.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Converting the date [$DateTime] to a universal sortable date time pattern based on the current culture [$($Host.CurrentCulture.Name)]."
                return [System.DateTime]::Parse($DateTime.TrimEnd('Z'), $Host.CurrentCulture).ToString([System.Globalization.DateTimeFormatInfo]::CurrentInfo.UniversalSortableDateTimePattern)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "The specified date/time [$DateTime] is not in a format recognized by the current culture [$($Host.CurrentCulture.Name)]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTUserProfiles
#
#-----------------------------------------------------------------------------

function Get-ADTUserProfiles
{
    <#
    .SYNOPSIS
        Get the User Profile Path, User Account SID, and the User Account Name for all users that log onto the machine and also the Default User.

    .DESCRIPTION
        Get the User Profile Path, User Account SID, and the User Account Name for all users that log onto the machine and also the Default User (which does not log on).
        Please note that the NTAccount property may be empty for some user profiles but the SID and ProfilePath properties will always be populated.

    .PARAMETER ExcludeNTAccount
        Specify NT account names in DOMAIN\username format to exclude from the list of user profiles.

    .PARAMETER IncludeSystemProfiles
        Include system profiles: SYSTEM, LOCAL SERVICE, NETWORK SERVICE.

    .PARAMETER IncludeServiceProfiles
        Include service (NT SERVICE) profiles.

    .PARAMETER IncludeIISAppPoolProfiles
        Include IIS AppPool profiles. Excluded by default as they don't parse well.

    .PARAMETER ExcludeDefaultUser
        Exclude the Default User.

    .PARAMETER LoadProfilePaths
        Load additional profile paths for each user profile.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.UserProfile

        Returns a PSADT.Types.UserProfile object with the following properties:
        - NTAccount
        - SID
        - ProfilePath

    .EXAMPLE
        Get-ADTUserProfiles

        Return the following properties for each user profile on the system: NTAccount, SID, ProfilePath.

    .EXAMPLE
        Get-ADTUserProfiles -ExcludeNTAccount CONTOSO\Robot,CONTOSO\ntadmin

        Return the following properties for each user profile on the system, except for 'Robot' and 'ntadmin': NTAccount, SID, ProfilePath.

    .EXAMPLE
        [string[]]$ProfilePaths = Get-ADTUserProfiles | Select-Object -ExpandProperty ProfilePath

        Return the user profile path for each user on the system. This information can then be used to make modifications under the user profile on the filesystem.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ExcludeNTAccount', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([PSADT.Types.UserProfile])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Security.Principal.NTAccount[]]$ExcludeNTAccount,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeSystemProfiles,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeServiceProfiles,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeIISAppPoolProfiles,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeDefaultUser,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$LoadProfilePaths
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $userProfileListRegKey = 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $excludedSids = "^S-1-5-($([System.String]::Join('|', $(
            if (!$IncludeSystemProfiles)
            {
                18  # System (or LocalSystem)
                19  # NT Authority (LocalService)
                20  # Network Service
            }
            if (!$IncludeServiceProfiles)
            {
                80  # NT Service
            }
            if (!$IncludeIISAppPoolProfiles)
            {
                82  # IIS AppPool
            }
        ))))"
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Getting the User Profile Path, User Account SID, and the User Account Name for all users that log onto the machine.'
        try
        {
            try
            {
                # Get the User Profile Path, User Account SID, and the User Account Name for all users that log onto the machine.
                & $Script:CommandTable.'Get-ItemProperty' -Path "$userProfileListRegKey\*" | & {
                    process
                    {
                        # Return early if the SID is to be excluded.
                        if ($_.PSChildName -match $excludedSids)
                        {
                            return
                        }

                        # Return early for accounts that have a null NTAccount.
                        if (!($ntAccount = & $Script:CommandTable.'ConvertTo-ADTNTAccountOrSID' -SID $_.PSChildName))
                        {
                            return
                        }

                        # Return early for excluded accounts.
                        if ($ExcludeNTAccount -contains $ntAccount)
                        {
                            return
                        }

                        # Establish base profile.
                        $userProfile = [PSADT.Types.UserProfile]::new(
                            $ntAccount,
                            $_.PSChildName,
                            $_.ProfileImagePath
                        )

                        # Append additional info if requested.
                        if ($LoadProfilePaths)
                        {
                            $userProfile = & $Script:CommandTable.'Invoke-ADTAllUsersRegistryAction' -UserProfiles $userProfile -ScriptBlock {
                                [PSADT.Types.UserProfile]::new(
                                    $_.NTAccount,
                                    $_.SID,
                                    $_.ProfilePath,
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'AppData' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Local AppData' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Desktop' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Personal' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Start Menu' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Environment' -Name 'TEMP' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Environment' -Name 'OneDrive' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath),
                                    $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Environment' -Name 'OneDriveCommercial' -SID $_.SID -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $_.ProfilePath)
                                )
                            }
                        }

                        # Write out the object to the pipeline.
                        return $userProfile
                    }
                }

                # Create a custom object for the Default User profile. Since the Default User is not an actual user account, it does not have a username or a SID.
                # We will make up a SID and add it to the custom object so that we have a location to load the default registry hive into later on.
                if (!$ExcludeDefaultUser)
                {
                    # The path to the default profile is stored in the default string value for the key.
                    $defaultUserProfilePath = (& $Script:CommandTable.'Get-ItemProperty' -LiteralPath $userProfileListRegKey).Default

                    # Retrieve additional information if requested.
                    if ($LoadProfilePaths)
                    {
                        return [PSADT.Types.UserProfile]::new(
                            'Default',
                            [System.Security.Principal.SecurityIdentifier]::new([System.Security.Principal.WellKnownSidType]::NullSid, $null),
                            $defaultUserProfilePath,
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'AppData' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Local AppData' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Desktop' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Personal' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Start Menu' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Environment' -Name 'TEMP' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Environment' -Name 'OneDrive' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath),
                            $((& $Script:CommandTable.'Get-ADTRegistryKey' -Key 'Microsoft.PowerShell.Core\Registry::HKEY_USERS\.DEFAULT\Environment' -Name 'OneDriveCommercial' -DoNotExpandEnvironmentNames) -replace '%USERPROFILE%', $defaultUserProfilePath)
                        )
                    }
                    return [PSADT.Types.UserProfile]::new(
                        'Default',
                        'S-1-0-0',
                        $defaultUserProfilePath
                    )
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Get-ADTWindowTitle
#
#-----------------------------------------------------------------------------

function Get-ADTWindowTitle
{
    <#
    .SYNOPSIS
        Search for an open window title and return details about the window.

    .DESCRIPTION
        Search for a window title. If window title searched for returns more than one result, then details for each window will be displayed.

        Returns the following properties for each window:
        - WindowTitle
        - WindowHandle
        - ParentProcess
        - ParentProcessMainWindowHandle
        - ParentProcessId

        Function does not work in SYSTEM context unless launched with "psexec.exe -s -i" to run it as an interactive process under the SYSTEM account.

    .PARAMETER WindowTitle
        One or more titles of the application window to search for using regex matching.

    .PARAMETER WindowHandle
        One or more window handles of the application window to search for.

    .PARAMETER ParentProcess
        One or more process names of the application window to search for.

    .PARAMETER GetAllWindowTitles
        Get titles for all open windows on the system.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.WindowInfo

        Returns a PSADT.Types.WindowInfo object with the following properties:
        - WindowTitle
        - WindowHandle
        - ParentProcess
        - ParentProcessMainWindowHandle
        - ParentProcessId

    .EXAMPLE
        Get-ADTWindowTitle -WindowTitle 'Microsoft Word'

        Gets details for each window that has the words "Microsoft Word" in the title.

    .EXAMPLE
        Get-ADTWindowTitle -GetAllWindowTitles

        Gets details for all windows with a title.

    .EXAMPLE
        Get-ADTWindowTitle -GetAllWindowTitles | Where-Object { $_.ParentProcess -eq 'WINWORD' }

        Get details for all windows belonging to Microsoft Word process with name "WINWORD".

    .NOTES
        An active ADT session is NOT required to use this function.

        Function does not work in SYSTEM context unless launched with "psexec.exe -s -i" to run it as an interactive process under the SYSTEM account.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'WindowTitle', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'WindowHandle', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ParentProcess', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'GetAllWindowTitles', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    [OutputType([PSADT.Types.WindowInfo])]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'SearchWinTitle')]
        [AllowEmptyString()]
        [System.String[]]$WindowTitle,

        [Parameter(Mandatory = $true, ParameterSetName = 'SearchWinHandle')]
        [AllowEmptyString()]
        [System.IntPtr[]]$WindowHandle,

        [Parameter(Mandatory = $true, ParameterSetName = 'SearchParentProcess')]
        [AllowEmptyString()]
        [System.String[]]$ParentProcess,

        [Parameter(Mandatory = $true, ParameterSetName = 'GetAllWinTitles')]
        [System.Management.Automation.SwitchParameter]$GetAllWindowTitles
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        # Announce commencement.
        switch ($PSCmdlet.ParameterSetName)
        {
            GetAllWinTitles
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Finding all open window title(s).'
                break
            }
            SearchWinTitle
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Finding open windows matching the specified title(s)."
                break
            }
            SearchWinHandle
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Finding open windows matching the specified handle(s)."
                break
            }
            SearchWinHandle
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Finding open windows matching the specified parent process(es)."
                break
            }
        }

        try
        {
            try
            {
                # Cache all running processes.
                $processes = [System.Diagnostics.Process]::GetProcesses() | & {
                    process
                    {
                        if ($WindowHandle -and ($_.MainWindowHandle -notin $WindowHandle))
                        {
                            return
                        }
                        if ($ParentProcess -and ($_.ProcessName -notin $ParentProcess))
                        {
                            return
                        }
                        return $_
                    }
                }

                # Get all window handles for visible windows and loop through the visible ones.
                [PSADT.GUI.UiAutomation]::EnumWindows() | & {
                    process
                    {
                        # Return early if we're null.
                        if ($null -eq $_)
                        {
                            return
                        }

                        # Return early if window isn't visible.
                        if (![PSADT.LibraryInterfaces.User32]::IsWindowVisible($_))
                        {
                            return
                        }

                        # Return early if the window doesn't have any text.
                        if (!($VisibleWindowTitle = [PSADT.GUI.UiAutomation]::GetWindowText($_)))
                        {
                            return
                        }

                        # Return early if the visible window title doesn't match our filter.
                        if ($WindowTitle -and ($VisibleWindowTitle -notmatch "($([System.String]::Join('|', $WindowTitle)))"))
                        {
                            return
                        }

                        # Return early if the window doesn't have an associated process.
                        if (!($process = $processes | & $Script:CommandTable.'Where-Object' -Property Id -EQ -Value ([PSADT.GUI.UiAutomation]::GetWindowThreadProcessId($_)) | & $Script:CommandTable.'Select-Object' -First 1))
                        {
                            return
                        }

                        # Build custom object with details about the window and the process.
                        return [PSADT.Types.WindowInfo]::new(
                            $VisibleWindowTitle,
                            $_,
                            $Process.ProcessName,
                            $Process.MainWindowHandle,
                            $Process.Id
                        )
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to get requested window title(s)."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Initialize-ADTFunction
#
#-----------------------------------------------------------------------------

function Initialize-ADTFunction
{
    <#
    .SYNOPSIS
        Initializes the ADT function environment.

    .DESCRIPTION
        Initializes the ADT function environment by setting up necessary variables and logging function start details. It ensures that the function always stops on errors and handles verbose logging.

    .PARAMETER Cmdlet
        The cmdlet that is being initialized.

    .PARAMETER SessionState
        The session state of the cmdlet.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Initialize-ADTFunction -Cmdlet $PSCmdlet

        Initializes the ADT function environment for the given cmdlet.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCmdlet]$Cmdlet,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SessionState]$SessionState
    )

    # Internal worker function to set variables within the caller's scope.
    function Set-CallerVariable
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This is an internal worker function that requires no end user confirmation.')]
        [CmdletBinding(SupportsShouldProcess = $false)]
        param
        (
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.String]$Name,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.Object]$Value
        )

        # Directly go up the scope tree if its an in-session function.
        if ($SessionState.Equals($ExecutionContext.SessionState))
        {
            & $Script:CommandTable.'Set-Variable' -Name $Name -Value $Value -Scope 2 -Force -Confirm:$false -WhatIf:$false
        }
        else
        {
            $SessionState.PSVariable.Set($Name, $Value)
        }
    }

    # Ensure this function always stops, no matter what.
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

    # Write debug log messages.
    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Function Start' -Source $Cmdlet.MyInvocation.MyCommand.Name -DebugMessage
    if ($Cmdlet.MyInvocation.BoundParameters.Count)
    {
        $CmdletBoundParameters = $Cmdlet.MyInvocation.BoundParameters | & $Script:CommandTable.'Format-Table' -Property @{ Label = 'Parameter'; Expression = { "[-$($_.Key)]" } }, @{ Label = 'Value'; Expression = { $_.Value }; Alignment = 'Left' }, @{ Label = 'Type'; Expression = { if ($_.Value) { $_.Value.GetType().Name } }; Alignment = 'Left' } -AutoSize -Wrap | & $Script:CommandTable.'Out-String' -Width ([System.Int32]::MaxValue)
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Function invoked with bound parameter(s):`n$CmdletBoundParameters" -Source $Cmdlet.MyInvocation.MyCommand.Name -DebugMessage
    }
    else
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Function invoked without any bound parameters.' -Source $Cmdlet.MyInvocation.MyCommand.Name -DebugMessage
    }

    # Amend the caller's $ErrorActionPreference to archive off their provided value so we can always stop on a dime.
    # For the caller-provided values, we deliberately use a string value to escape issues when 'Ignore' is passed.
    # https://github.com/PowerShell/PowerShell/issues/1759#issuecomment-442916350
    if ($Cmdlet.MyInvocation.BoundParameters.ContainsKey('ErrorAction'))
    {
        # Caller's value directly against the function.
        Set-CallerVariable -Name OriginalErrorAction -Value $Cmdlet.MyInvocation.BoundParameters.ErrorAction.ToString()
    }
    elseif ($PSBoundParameters.ContainsKey('ErrorAction'))
    {
        # A function's own specified override.
        Set-CallerVariable -Name OriginalErrorAction -Value $PSBoundParameters.ErrorAction.ToString()
    }
    else
    {
        # The module's default ErrorActionPreference.
        Set-CallerVariable -Name OriginalErrorAction -Value $Script:ErrorActionPreference
    }
    Set-CallerVariable -Name ErrorActionPreference -Value $Script:ErrorActionPreference
}


#-----------------------------------------------------------------------------
#
# MARK: Initialize-ADTModule
#
#-----------------------------------------------------------------------------

function Initialize-ADTModule
{
    <#
    .SYNOPSIS
        Initializes the ADT module by setting up necessary configurations and environment.

    .DESCRIPTION
        The Initialize-ADTModule function sets up the environment for the ADT module by initializing necessary variables, configurations, and string tables. It ensures that the module is not initialized while there is an active ADT session in progress. This function prepares the module for use by clearing callbacks, sessions, and setting up the environment table.

    .PARAMETER ScriptDirectory
        An override directory to use for config and string loading.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Initialize-ADTModule

        Initializes the ADT module with the default settings and configurations.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ScriptDirectory -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ScriptDirectory -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String[]]$ScriptDirectory
    )

    begin
    {
        # Log our start time to clock the module init duration.
        $moduleInitStart = [System.DateTime]::Now

        # Ensure this function isn't being called mid-flight.
        if (& $Script:CommandTable.'Test-ADTSessionActive')
        {
            $naerParams = @{
                Exception = [System.InvalidOperationException]::new("This function cannot be called while there is an active ADTSession in progress.")
                Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                ErrorId = 'InitWithActiveSessionError'
                TargetObject = & $Script:CommandTable.'Get-ADTSession'
                RecommendedAction = "Please attempt module re-initialization once the active ADTSession(s) have been closed."
            }
            $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
        }
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Specify the base directory used when searching for config and string tables.
                $Script:ADT.Directories.Script = if ($PSBoundParameters.ContainsKey('ScriptDirectory'))
                {
                    $ScriptDirectory
                }
                else
                {
                    $Script:ADT.Directories.Defaults.Script
                }

                # Initialize remaining directory paths.
                'Config', 'Strings' | & {
                    process
                    {
                        [System.String[]]$Script:ADT.Directories.$_ = foreach ($directory in $Script:ADT.Directories.Script)
                        {
                            if ([System.IO.File]::Exists([System.IO.Path]::Combine($directory, $_, "$($_.ToLower()).psd1")))
                            {
                                [System.IO.Path]::Combine($directory, $_)
                            }
                        }
                        if ($null -eq $Script:ADT.Directories.$_)
                        {
                            [System.String[]]$Script:ADT.Directories.$_ = $Script:ADT.Directories.Defaults.$_
                        }
                    }
                }

                # Initialize the module's global state.
                $Script:ADT.Environment = & $Script:CommandTable.'New-ADTEnvironmentTable'
                $Script:ADT.Config = & $Script:CommandTable.'Import-ADTConfig' -BaseDirectory $Script:ADT.Directories.Config
                $Script:ADT.Language = & $Script:CommandTable.'Get-ADTStringLanguage'
                $Script:ADT.Strings = & $Script:CommandTable.'Import-ADTModuleDataFile' -BaseDirectory $Script:ADT.Directories.Strings -FileName strings.psd1 -UICulture $Script:ADT.Language -IgnorePolicy
                $Script:ADT.Sessions.Clear()
                $Script:ADT.TerminalServerMode = $false
                $Script:ADT.LastExitCode = 0

                # Calculate how long this process took before finishing.
                $Script:ADT.Durations.ModuleInit = [System.DateTime]::Now - $moduleInitStart
                $Script:ADT.Initialized = $true
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Install-ADTMSUpdates
#
#-----------------------------------------------------------------------------

function Install-ADTMSUpdates
{
    <#
    .SYNOPSIS
        Install all Microsoft Updates in a given directory.

    .DESCRIPTION
        Install all Microsoft Updates of type ".exe", ".msu", or ".msp" in a given directory (recursively search directory). The function will check if the update is already installed and skip it if it is. It handles older redistributables and different types of updates appropriately.

    .PARAMETER Directory
        Directory containing the updates.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Install-ADTMSUpdates -Directory "$($adtSession.DirFiles)\MSUpdates"

        Installs all Microsoft Updates found in the specified directory.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Directory
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $kbPattern = '(?i)kb\d{6,8}'
    }

    process
    {
        # Get all hotfixes and install if required.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Recursively installing all Microsoft Updates in directory [$Directory]."
        foreach ($file in (& $Script:CommandTable.'Get-ChildItem' -LiteralPath $Directory -Recurse -Include ('*.exe', '*.msu', '*.msp')))
        {
            try
            {
                try
                {
                    if ($file.Name -match 'redist')
                    {
                        # Handle older redistributables (ie, VC++ 2005)
                        [System.Version]$redistVersion = $file.VersionInfo.ProductVersion
                        [System.String]$redistDescription = $file.VersionInfo.FileDescription
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Installing [$redistDescription $redistVersion]..."
                        if ($redistDescription -match 'Win32 Cabinet Self-Extractor')
                        {
                            & $Script:CommandTable.'Start-ADTProcess' -FilePath $file.FullName -ArgumentList '/q' -WindowStyle 'Hidden' -IgnoreExitCodes '*'
                        }
                        else
                        {
                            & $Script:CommandTable.'Start-ADTProcess' -FilePath $file.FullName -ArgumentList '/quiet /norestart' -WindowStyle 'Hidden' -IgnoreExitCodes '*'
                        }
                    }
                    elseif ($kbNumber = [System.Text.RegularExpressions.Regex]::Match($file.Name, $kbPattern).ToString())
                    {
                        # Check to see whether the KB is already installed
                        if (& $Script:CommandTable.'Test-ADTMSUpdates' -KbNumber $kbNumber)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "KB Number [$kbNumber] is already installed. Continue..."
                            continue
                        }
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "KB Number [$KBNumber] was not detected and will be installed."
                        switch ($file.Extension)
                        {
                            '.exe'
                            {
                                # Installation type for executables (i.e., Microsoft Office Updates).
                                & $Script:CommandTable.'Start-ADTProcess' -FilePath $file.FullName -ArgumentList '/quiet /norestart' -WindowStyle 'Hidden' -IgnoreExitCodes '*'
                                break
                            }
                            '.msu'
                            {
                                # Installation type for Windows updates using Windows Update Standalone Installer.
                                & $Script:CommandTable.'Start-ADTProcess' -FilePath "$([System.Environment]::SystemDirectory)\wusa.exe" -ArgumentList "`"$($file.FullName)`" /quiet /norestart" -WindowStyle 'Hidden' -IgnoreExitCodes '*'
                                break
                            }
                            '.msp'
                            {
                                # Installation type for Windows Installer Patch
                                & $Script:CommandTable.'Start-ADTMsiProcess' -Action 'Patch' -Path $file.FullName -IgnoreExitCodes '*'
                                break
                            }
                        }
                    }
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Install-ADTSCCMSoftwareUpdates
#
#-----------------------------------------------------------------------------

function Install-ADTSCCMSoftwareUpdates
{
    <#
    .SYNOPSIS
        Scans for outstanding SCCM updates to be installed and installs the pending updates.

    .DESCRIPTION
        Scans for outstanding SCCM updates to be installed and installs the pending updates.
        Only compatible with SCCM 2012 Client or higher. This function can take several minutes to run.

    .PARAMETER SoftwareUpdatesScanWaitInSeconds
        The amount of time to wait in seconds for the software updates scan to complete. Default is: 180 seconds.

    .PARAMETER WaitForPendingUpdatesTimeout
        The amount of time to wait for missing and pending updates to install before exiting the function. Default is: 45 minutes.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Install-ADTSCCMSoftwareUpdates

        Scans for outstanding SCCM updates and installs the pending updates with default wait times.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$SoftwareUpdatesScanWaitInSeconds = 180,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$WaitForPendingUpdatesTimeout = [System.TimeSpan]::FromMinutes(45)
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        try
        {
            try
            {
                # If SCCM 2007 Client or lower, exit function.
                if (($SCCMClientVersion = & $Script:CommandTable.'Get-ADTSCCMClientVersion').Major -le 4)
                {
                    $naerParams = @{
                        Exception = [System.Data.VersionNotFoundException]::new('SCCM 2007 or lower, which is incompatible with this function, was detected on this system.')
                        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                        ErrorId = 'CcmExecVersionLowerThanMinimum'
                        TargetObject = $SCCMClientVersion
                        RecommendedAction = "Please review the installed CcmExec client and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Trigger SCCM client scan for Software Updates.
                $StartTime = [System.DateTime]::Now
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Triggering SCCM client scan for Software Updates...'
                & $Script:CommandTable.'Invoke-ADTSCCMTask' -ScheduleID 'SoftwareUpdatesScan'
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "The SCCM client scan for Software Updates has been triggered. The script is suspended for [$SoftwareUpdatesScanWaitInSeconds] seconds to let the update scan finish."
                & $Script:CommandTable.'Start-Sleep' -Seconds $SoftwareUpdatesScanWaitInSeconds

                # Find the number of missing updates.
                try
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Getting the number of missing updates...'
                    [Microsoft.Management.Infrastructure.CimInstance[]]$CMMissingUpdates = & $Script:CommandTable.'Get-CimInstance' -Namespace ROOT\CCM\ClientSDK -Query "SELECT * FROM CCM_SoftwareUpdate WHERE ComplianceState = '0'"
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to find the number of missing software updates.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 2
                    throw
                }

                # Install missing updates and wait for pending updates to finish installing.
                if (!$CMMissingUpdates.Count)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'There are no missing updates.'
                    return
                }

                # Install missing updates.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Installing missing updates. The number of missing updates is [$($CMMissingUpdates.Count)]."
                $null = & $Script:CommandTable.'Invoke-CimMethod' -Namespace ROOT\CCM\ClientSDK -ClassName CCM_SoftwareUpdatesManager -MethodName InstallUpdates -Arguments @{ CCMUpdates = $CMMissingUpdates }

                # Wait for pending updates to finish installing or the timeout value to expire.
                do
                {
                    & $Script:CommandTable.'Start-Sleep' -Seconds 60
                    [Microsoft.Management.Infrastructure.CimInstance[]]$CMInstallPendingUpdates = & $Script:CommandTable.'Get-CimInstance' -Namespace ROOT\CCM\ClientSDK -Query 'SELECT * FROM CCM_SoftwareUpdate WHERE EvaluationState = 6 or EvaluationState = 7'
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The number of updates pending installation is [$($CMInstallPendingUpdates.Count)]."
                }
                while (($CMInstallPendingUpdates.Count -ne 0) -and ([System.DateTime]::Now - $StartTime) -lt $WaitForPendingUpdatesTimeout)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to trigger installation of missing software updates."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTAllUsersRegistryAction
#
#-----------------------------------------------------------------------------

function Invoke-ADTAllUsersRegistryAction
{
    <#
    .SYNOPSIS
        Set current user registry settings for all current users and any new users in the future.

    .DESCRIPTION
        Set HKCU registry settings for all current and future users by loading their NTUSER.dat registry hive file, and making the modifications.

        This function will modify HKCU settings for all users even when executed under the SYSTEM account and can be used as an alternative to using ActiveSetup for registry settings.

        To ensure new users in the future get the registry edits, the Default User registry hive used to provision the registry for new users is modified.

        The advantage of using this function over ActiveSetup is that a user does not have to log off and log back on before the changes take effect.

    .PARAMETER ScriptBlock
        Script block which contains HKCU registry actions to be run for all users on the system.

    .PARAMETER UserProfiles
        Specify the user profiles to modify HKCU registry settings for. Default is all user profiles except for system profiles.

    .PARAMETER SkipUnloadedProfiles
        Specifies that unloaded registry hives should be skipped and not be loaded by the function.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Invoke-ADTAllUsersRegistryAction -ScriptBlock {
            Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'qmenable' -Value 0 -Type DWord -SID $_.SID
            Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'updatereliabilitydata' -Value 1 -Type DWord -SID $_.SID
        }

        Example demonstrating the setting of two values within each user's HKEY_CURRENT_USER hive.

    .EXAMPLE
        Invoke-ADTAllUsersRegistryAction {
            Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'qmenable' -Value 0 -Type DWord -SID $_.SID
            Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'updatereliabilitydata' -Value 1 -Type DWord -SID $_.SID
        }

        As the previous example, but showing how to use ScriptBlock as a positional parameter with no name specified.

    .EXAMPLE
        Invoke-ADTAllUsersRegistryAction -UserProfiles (Get-ADTUserProfiles -ExcludeDefaultUser) -ScriptBlock {
            Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'qmenable' -Value 0 -Type DWord -SID $_.SID
            Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Office\14.0\Common' -Name 'updatereliabilitydata' -Value 1 -Type DWord -SID $_.SID
        }

        As the previous example, but sending specific user profiles through to exclude the Default profile.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock[]]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.UserProfile[]]$UserProfiles = (& $Script:CommandTable.'Get-ADTUserProfiles'),

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipUnloadedProfiles
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Internal function to unload registry hives at the end of the operation.
        function Dismount-UserProfileRegistryHive
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unloading the User [$($UserProfile.NTAccount)] registry hive in path [HKEY_USERS\$($UserProfile.SID)]."
            $null = & "$([System.Environment]::SystemDirectory)\reg.exe" UNLOAD "HKEY_USERS\$($UserProfile.SID)" 2>&1
        }
    }

    process
    {
        foreach ($UserProfile in $UserProfiles)
        {
            $ManuallyLoadedRegHive = $false
            try
            {
                try
                {
                    # Set the path to the user's registry hive file.
                    $UserRegistryHiveFile = & $Script:CommandTable.'Join-Path' -Path $UserProfile.ProfilePath -ChildPath 'NTUSER.DAT'

                    # Load the User profile registry hive if it is not already loaded because the User is logged in.
                    if (!(& $Script:CommandTable.'Test-Path' -LiteralPath "Microsoft.PowerShell.Core\Registry::HKEY_USERS\$($UserProfile.SID)"))
                    {
                        # Only load the profile if we've been asked to.
                        if ($SkipUnloadedProfiles)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Skipping User [$($UserProfile.NTAccount)] as the registry hive is not loaded."
                            continue
                        }

                        # Load the User registry hive if the registry hive file exists.
                        if (![System.IO.File]::Exists($UserRegistryHiveFile))
                        {
                            $naerParams = @{
                                Exception = [System.IO.FileNotFoundException]::new("Failed to find the registry hive file [$UserRegistryHiveFile] for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)]. Continue...")
                                Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                                ErrorId = 'UserRegistryHiveFileNotFound'
                                TargetObject = $UserRegistryHiveFile
                                RecommendedAction = "Please confirm the state of this user profile and try again."
                            }
                            throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                        }

                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Loading the User [$($UserProfile.NTAccount)] registry hive in path [HKEY_USERS\$($UserProfile.SID)]."
                        $null = & "$([System.Environment]::SystemDirectory)\reg.exe" LOAD "HKEY_USERS\$($UserProfile.SID)" $UserRegistryHiveFile 2>&1
                        $ManuallyLoadedRegHive = $true
                    }

                    # Invoke changes against registry.
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Executing scriptblock to modify HKCU registry settings for all users.'
                    & $Script:CommandTable.'ForEach-Object' -InputObject $UserProfile -Begin $null -End $null -Process $ScriptBlock
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to modify the registry hive for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)]`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
            }
            finally
            {
                if ($ManuallyLoadedRegHive)
                {
                    try
                    {
                        try
                        {
                            Dismount-UserProfileRegistryHive
                        }
                        catch
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "REG.exe failed to unload the registry hive with exit code [$($Global:LASTEXITCODE)] and error message [$($_.Exception.Message)]." -Severity 2
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Performing manual garbage collection to ensure successful unloading of registry hive." -Severity 2
                            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); [System.Threading.Thread]::Sleep(5000)
                            Dismount-UserProfileRegistryHive
                        }
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to unload the registry hive for User [$($UserProfile.NTAccount)] with SID [$($UserProfile.SID)]. REG.exe exit code [$Global:LASTEXITCODE]. Error message: [$($_.Exception.Message)]" -Severity 3
                    }
                }
            }
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTCommandWithRetries
#
#-----------------------------------------------------------------------------

function Invoke-ADTCommandWithRetries
{
    <#
    .SYNOPSIS
        Drop-in replacement for any cmdlet/function where a retry is desirable due to transient issues.

    .DESCRIPTION
        This function invokes the specified cmdlet/function, accepting all of its parameters but retries an operation for the configured value before throwing.

    .PARAMETER Command
        The name of the command to invoke.

    .PARAMETER Retries
        How many retries to perform before throwing.

    .PARAMETER SleepSeconds
        How many seconds to sleep between retries.

    .PARAMETER Parameters
        A 'ValueFromRemainingArguments' parameter to collect the parameters as would be passed to the provided Command.

        While values can be directly provided to this parameter, it's not designed to be explicitly called.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Object

        Invoke-ADTCommandWithRetries returns the output of the invoked command.

    .EXAMPLE
        Invoke-ADTCommandWithRetries -Command Invoke-WebRequest -Uri https://aka.ms/getwinget -OutFile "$($adtSession.DirSupportFiles)\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"

        Downloads the latest WinGet installer to the SupportFiles directory.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Command,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$Retries = 3,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 60)]
        [System.UInt32]$SleepSeconds = 5,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$Parameters
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Attempt to get command from our lookup table.
                $commandObj = if ($Command -is [System.Management.Automation.CommandInfo])
                {
                    $Command
                }
                elseif ($Script:CommandTable.ContainsKey($Command))
                {
                    $Script:CommandTable.$Command
                }
                else
                {
                    & $Script:CommandTable.'Get-Command' -Name $Command
                }

                # Convert the passed parameters into a dictionary for splatting onto the command.
                $boundParams = & $Script:CommandTable.'Convert-ADTValuesFromRemainingArguments' -RemainingArguments $Parameters
                $callerName = (& $Script:CommandTable.'Get-PSCallStack')[1].Command

                # Perform the request, and retry it as per the configured values.
                for ($i = 0; $i -lt $Retries; $i++)
                {
                    try
                    {
                        return (& $commandObj @boundParams)
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "The invocation to '$($commandObj.Name)' failed with message: $($_.Exception.Message.TrimEnd('.')). Trying again in $SleepSeconds second$(if (!$SleepSeconds.Equals(1)) {'s'})." -Severity 2 -Source $callerName
                        [System.Threading.Thread]::Sleep($SleepSeconds * 1000)
                        $errorRecord = $_
                    }
                }

                # If we're here, we failed too many times. Throw the captured ErrorRecord.
                throw $errorRecord
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTFunctionErrorHandler
#
#-----------------------------------------------------------------------------

function Invoke-ADTFunctionErrorHandler
{
    <#
    .SYNOPSIS
        Handles errors within ADT functions by logging and optionally passing through the error.

    .DESCRIPTION
        This function handles errors within ADT functions by logging the error message and optionally passing through the error record. It recovers the true ErrorActionPreference set by the caller and sets it within the function. If a log message is provided, it appends the resolved error record to the log message. Depending on the ErrorActionPreference, it either throws a terminating error or writes a non-terminating error.

    .PARAMETER Cmdlet
        The cmdlet that is calling this function.

    .PARAMETER SessionState
        The session state of the calling cmdlet.

    .PARAMETER ErrorRecord
        The error record to handle.

    .PARAMETER LogMessage
        The error message to write to the active ADTSession's log file.

    .PARAMETER DisableErrorResolving
        If specified, the function will not append the resolved error record to the log message.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Invoke-ADTFunctionErrorHandler -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_

        Handles the error within the calling cmdlet and logs it.

    .EXAMPLE
        Invoke-ADTFunctionErrorHandler -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "An error occurred" -DisableErrorResolving

        Handles the error within the calling cmdlet, logs a custom message without resolving the error record, and logs it.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(DefaultParameterSetName = 'None')]
    [OutputType([System.Management.Automation.ErrorRecord])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCmdlet]$Cmdlet,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SessionState]$SessionState,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory = $true, ParameterSetName = 'LogMessage')]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogMessage,

        [Parameter(Mandatory = $false, ParameterSetName = 'LogMessage')]
        [System.Management.Automation.SwitchParameter]$DisableErrorResolving
    )

    # Recover true ErrorActionPreference the caller may have set,
    # unless an ErrorAction was specifically provided to this function.
    $ErrorActionPreference = if ($PSBoundParameters.ContainsKey('ErrorAction'))
    {
        $PSBoundParameters.ErrorAction
    }
    elseif ($SessionState.Equals($ExecutionContext.SessionState))
    {
        & $Script:CommandTable.'Get-Variable' -Name OriginalErrorAction -Scope 1 -ValueOnly
    }
    else
    {
        $SessionState.PSVariable.Get('OriginalErrorAction').Value
    }

    # If the caller hasn't specified a LogMessage, use the ErrorRecord's message.
    if ([System.String]::IsNullOrWhiteSpace($LogMessage))
    {
        $LogMessage = $ErrorRecord.Exception.Message
    }

    # Write-Error enforces its own name against the Activity, let's re-write it.
    if ($ErrorRecord.CategoryInfo.Activity -match '^Write-Error$')
    {
        $ErrorRecord.CategoryInfo.Activity = $Cmdlet.MyInvocation.MyCommand.Name
    }

    # Write out the error to the log file.
    if (!$DisableErrorResolving)
    {
        $LogMessage += "`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $ErrorRecord)"
    }
    & $Script:CommandTable.'Write-ADTLogEntry' -Message $LogMessage -Source $Cmdlet.MyInvocation.MyCommand.Name -Severity 3

    # If we're stopping, throw a terminating error. While WriteError will terminate if stopping,
    # this can also write out an [System.Management.Automation.ActionPreferenceStopException] object.
    if ($ErrorActionPreference.Equals([System.Management.Automation.ActionPreference]::Stop))
    {
        $Cmdlet.ThrowTerminatingError($ErrorRecord)
    }
    elseif (!(& $Script:CommandTable.'Test-ADTSessionActive') -or ($ErrorActionPreference -notmatch '^(SilentlyContinue|Ignore)$'))
    {
        $Cmdlet.WriteError($ErrorRecord)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTObjectMethod
#
#-----------------------------------------------------------------------------

function Invoke-ADTObjectMethod
{
    <#
    .SYNOPSIS
        Invoke method on any object.

    .DESCRIPTION
        Invoke method on any object with or without using named parameters.

    .PARAMETER InputObject
        Specifies an object which has methods that can be invoked.

    .PARAMETER MethodName
        Specifies the name of a method to invoke.

    .PARAMETER ArgumentList
        Argument to pass to the method being executed. Allows execution of method without specifying named parameters.

    .PARAMETER Parameter
        Argument to pass to the method being executed. Allows execution of method by using named parameters.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Object

        The object returned by the method being invoked.

    .EXAMPLE
        PS C:\>$ShellApp = New-Object -ComObject 'Shell.Application'
        PS C:\>$null = Invoke-ADTObjectMethod -InputObject $ShellApp -MethodName 'MinimizeAll'

        Minimizes all windows.

    .EXAMPLE
        PS C:\>$ShellApp = New-Object -ComObject 'Shell.Application'
        PS C:\>$null = Invoke-ADTObjectMethod -InputObject $ShellApp -MethodName 'Explore' -Parameter @{'vDir'='C:\Windows'}

        Opens the C:\Windows folder in a Windows Explorer window.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(DefaultParameterSetName = 'Positional')]
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

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                switch ($PSCmdlet.ParameterSetName)
                {
                    Named
                    {
                        # Invoke method by using parameter names.
                        return $InputObject.GetType().InvokeMember($MethodName, [System.Reflection.BindingFlags]::InvokeMethod, $null, $InputObject, ([System.Object[]]$Parameter.Values), $null, $null, ([System.String[]]$Parameter.Keys))
                    }
                    Positional
                    {
                        # Invoke method without using parameter names.
                        return $InputObject.GetType().InvokeMember($MethodName, [System.Reflection.BindingFlags]::InvokeMethod, $null, $InputObject, $ArgumentList, $null, $null, $null)
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTRegSvr32
#
#-----------------------------------------------------------------------------

function Invoke-ADTRegSvr32
{
    <#
    .SYNOPSIS
        Register or unregister a DLL file.

    .DESCRIPTION
        Register or unregister a DLL file using regsvr32.exe. This function determines the bitness of the DLL file and uses the appropriate version of regsvr32.exe to perform the action. It supports both 32-bit and 64-bit DLL files on corresponding operating systems.

    .PARAMETER FilePath
        Path to the DLL file.

    .PARAMETER Action
        Specify whether to register or unregister the DLL.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return objects.

    .EXAMPLE
        Invoke-ADTRegSvr32 -FilePath "C:\Test\DcTLSFileToDMSComp.dll" -Action 'Register'

        Registers the specified DLL file.

    .EXAMPLE
        Invoke-ADTRegSvr32 -FilePath "C:\Test\DcTLSFileToDMSComp.dll" -Action 'Unregister'

        Unregisters the specified DLL file.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_) -and ([System.IO.Path]::GetExtension($_) -ne '.dll'))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist or is not a DLL file.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Register', 'Unregister')]
        [System.String]$Action
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Define parameters to pass to regsrv32.exe.
        $ActionParameters = switch ($Action = $Host.CurrentCulture.TextInfo.ToTitleCase($Action.ToLower()))
        {
            Register
            {
                "/s `"$FilePath`""
                break
            }
            Unregister
            {
                "/s /u `"$FilePath`""
                break
            }
        }
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "$Action DLL file [$FilePath]."
        try
        {
            try
            {
                # Determine the bitness of the DLL file.
                if ((($DLLFileBitness = & $Script:CommandTable.'Get-ADTPEFileArchitecture' -FilePath $FilePath) -ne [PSADT.Shared.SystemArchitecture]::AMD64) -and ($DLLFileBitness -ne [PSADT.Shared.SystemArchitecture]::i386))
                {
                    $naerParams = @{
                        Exception = [System.PlatformNotSupportedException]::new("File [$filePath] has a detected file architecture of [$DLLFileBitness]. Only 32-bit or 64-bit DLL files can be $($Action.ToLower() + 'ed').")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                        ErrorId = 'DllFileArchitectureError'
                        TargetObject = $FilePath
                        RecommendedAction = "Please review the supplied DLL FilePath and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Get the correct path to regsrv32.exe for the system and DLL file.
                $RegSvr32Path = if ([System.Environment]::Is64BitOperatingSystem)
                {
                    if ($DLLFileBitness -eq [PSADT.Shared.SystemArchitecture]::AMD64)
                    {
                        if ([System.Environment]::Is64BitProcess)
                        {
                            "$([System.Environment]::SystemDirectory)\regsvr32.exe"
                        }
                        else
                        {
                            "$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows))\sysnative\regsvr32.exe"
                        }
                    }
                    elseif ($DLLFileBitness -eq [PSADT.Shared.SystemArchitecture]::i386)
                    {
                        "$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::SystemX86))\regsvr32.exe"
                    }
                }
                elseif ($DLLFileBitness -eq [PSADT.Shared.SystemArchitecture]::i386)
                {
                    "$([System.Environment]::SystemDirectory)\regsvr32.exe"
                }
                else
                {
                    $naerParams = @{
                        Exception = [System.PlatformNotSupportedException]::new("File [$filePath] cannot be $($Action.ToLower()) because it is a 64-bit file on a 32-bit operating system.")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                        ErrorId = 'DllFileArchitectureError'
                        TargetObject = $FilePath
                        RecommendedAction = "Please review the supplied DLL FilePath and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Register the DLL file and measure the success.
                if (($ExecuteResult = & $Script:CommandTable.'Start-ADTProcess' -FilePath $RegSvr32Path -ArgumentList $ActionParameters -WindowStyle Hidden -PassThru).ExitCode -ne 0)
                {
                    if ($ExecuteResult.ExitCode -eq 60002)
                    {
                        $naerParams = @{
                            Exception = [System.InvalidOperationException]::new("Start-ADTProcess function failed with exit code [$($ExecuteResult.ExitCode)].")
                            Category = [System.Management.Automation.ErrorCategory]::OperationStopped
                            ErrorId = 'ProcessInvocationError'
                            TargetObject = "$FilePath $ActionParameters"
                            RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                    else
                    {
                        $naerParams = @{
                            Exception = [System.InvalidOperationException]::new("regsvr32.exe failed with exit code [$($ExecuteResult.ExitCode)].")
                            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                            ErrorId = 'ProcessInvocationError'
                            TargetObject = "$FilePath $ActionParameters"
                            RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to $($Action.ToLower()) DLL file."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Invoke-ADTSCCMTask
#
#-----------------------------------------------------------------------------

function Invoke-ADTSCCMTask
{
    <#
    .SYNOPSIS
        Triggers SCCM to invoke the requested schedule task ID.

    .DESCRIPTION
        Triggers SCCM to invoke the requested schedule task ID. This function supports a variety of schedule IDs compatible with different versions of the SCCM client. It ensures that the correct schedule IDs are used based on the SCCM client version.

    .PARAMETER ScheduleId
        Name of the schedule id to trigger.

        Options: HardwareInventory, SoftwareInventory, HeartbeatDiscovery, SoftwareInventoryFileCollection, RequestMachinePolicy, EvaluateMachinePolicy, LocationServicesCleanup, SoftwareMeteringReport, SourceUpdate, PolicyAgentCleanup, RequestMachinePolicy2, CertificateMaintenance, PeerDistributionPointStatus, PeerDistributionPointProvisioning, ComplianceIntervalEnforcement, SoftwareUpdatesAgentAssignmentEvaluation, UploadStateMessage, StateMessageManager, SoftwareUpdatesScan, AMTProvisionCycle, UpdateStorePolicy, StateSystemBulkSend, ApplicationManagerPolicyAction, PowerManagementStartSummarizer

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Invoke-ADTSCCMTask -ScheduleId 'SoftwareUpdatesScan'

        Triggers the 'SoftwareUpdatesScan' schedule task in SCCM.

    .EXAMPLE
        Invoke-ADTSCCMTask -ScheduleId 'HardwareInventory'

        Triggers the 'HardwareInventory' schedule task in SCCM.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('HardwareInventory', 'SoftwareInventory', 'HeartbeatDiscovery', 'SoftwareInventoryFileCollection', 'RequestMachinePolicy', 'EvaluateMachinePolicy', 'LocationServicesCleanup', 'SoftwareMeteringReport', 'SourceUpdate', 'PolicyAgentCleanup', 'RequestMachinePolicy2', 'CertificateMaintenance', 'PeerDistributionPointStatus', 'PeerDistributionPointProvisioning', 'ComplianceIntervalEnforcement', 'SoftwareUpdatesAgentAssignmentEvaluation', 'UploadStateMessage', 'StateMessageManager', 'SoftwareUpdatesScan', 'AMTProvisionCycle', 'UpdateStorePolicy', 'StateSystemBulkSend', 'ApplicationManagerPolicyAction', 'PowerManagementStartSummarizer')]
        [System.String]$ScheduleID
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Create a hashtable of Schedule IDs compatible with SCCM Client 2007.
        $ScheduleIds = @{
            HardwareInventory = '{00000000-0000-0000-0000-000000000001}'  # Hardware Inventory Collection Task
            SoftwareInventory = '{00000000-0000-0000-0000-000000000002}'  # Software Inventory Collection Task
            HeartbeatDiscovery = '{00000000-0000-0000-0000-000000000003}'  # Heartbeat Discovery Cycle
            SoftwareInventoryFileCollection = '{00000000-0000-0000-0000-000000000010}'  # Software Inventory File Collection Task
            RequestMachinePolicy = '{00000000-0000-0000-0000-000000000021}'  # Request Machine Policy Assignments
            EvaluateMachinePolicy = '{00000000-0000-0000-0000-000000000022}'  # Evaluate Machine Policy Assignments
            RefreshDefaultMp = '{00000000-0000-0000-0000-000000000023}'  # Refresh Default MP Task
            RefreshLocationServices = '{00000000-0000-0000-0000-000000000024}'  # Refresh Location Services Task
            LocationServicesCleanup = '{00000000-0000-0000-0000-000000000025}'  # Location Services Cleanup Task
            SoftwareMeteringReport = '{00000000-0000-0000-0000-000000000031}'  # Software Metering Report Cycle
            SourceUpdate = '{00000000-0000-0000-0000-000000000032}'  # Source Update Manage Update Cycle
            PolicyAgentCleanup = '{00000000-0000-0000-0000-000000000040}'  # Policy Agent Cleanup Cycle
            RequestMachinePolicy2 = '{00000000-0000-0000-0000-000000000042}'  # Request Machine Policy Assignments
            CertificateMaintenance = '{00000000-0000-0000-0000-000000000051}'  # Certificate Maintenance Cycle
            PeerDistributionPointStatus = '{00000000-0000-0000-0000-000000000061}'  # Peer Distribution Point Status Task
            PeerDistributionPointProvisioning = '{00000000-0000-0000-0000-000000000062}'  # Peer Distribution Point Provisioning Status Task
            ComplianceIntervalEnforcement = '{00000000-0000-0000-0000-000000000071}'  # Compliance Interval Enforcement
            SoftwareUpdatesAgentAssignmentEvaluation = '{00000000-0000-0000-0000-000000000108}'  # Software Updates Agent Assignment Evaluation Cycle
            UploadStateMessage = '{00000000-0000-0000-0000-000000000111}'  # Send Unsent State Messages
            StateMessageManager = '{00000000-0000-0000-0000-000000000112}'  # State Message Manager Task
            SoftwareUpdatesScan = '{00000000-0000-0000-0000-000000000113}'  # Force Update Scan
            AMTProvisionCycle = '{00000000-0000-0000-0000-000000000120}'  # AMT Provision Cycle
        }
    }

    process
    {
        try
        {
            try
            {
                # If SCCM 2012 Client or higher, modify hashtabe containing Schedule IDs so that it only has the ones compatible with this version of the SCCM client.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoke SCCM Schedule Task ID [$ScheduleId]..."
                if ((& $Script:CommandTable.'Get-ADTSCCMClientVersion').Major -ge 5)
                {
                    $ScheduleIds.Remove('PeerDistributionPointStatus')
                    $ScheduleIds.Remove('PeerDistributionPointProvisioning')
                    $ScheduleIds.Remove('ComplianceIntervalEnforcement')
                    $ScheduleIds.Add('UpdateStorePolicy', '{00000000-0000-0000-0000-000000000114}') # Update Store Policy
                    $ScheduleIds.Add('StateSystemBulkSend', '{00000000-0000-0000-0000-000000000116}') # State System Policy Bulk Send Low
                    $ScheduleIds.Add('ApplicationManagerPolicyAction', '{00000000-0000-0000-0000-000000000121}') # Application Manager Policy Action
                    $ScheduleIds.Add('PowerManagementStartSummarizer', '{00000000-0000-0000-0000-000000000131}') # Power Management Start Summarizer
                }

                # Determine if the requested Schedule ID is available on this version of the SCCM Client.
                if (!$ScheduleIds.ContainsKey($ScheduleId))
                {
                    $naerParams = @{
                        Exception = [System.ApplicationException]::new("The requested ScheduleId [$ScheduleId] is not available with this version of the SCCM Client [$SCCMClientVersion].")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidData
                        ErrorId = 'CcmExecInvalidScheduleId'
                        RecommendedAction = 'Please check the supplied ScheduleId and try again.'
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Trigger SCCM task.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Triggering SCCM Task ID [$ScheduleId]."
                $null = & $Script:CommandTable.'Invoke-CimMethod' -Namespace ROOT\CCM -ClassName SMS_Client -MethodName TriggerSchedule -Arguments @{ sScheduleID = $ScheduleIds.$ScheduleID }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to trigger SCCM Schedule Task ID [$($ScheduleIds.$ScheduleId)]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Mount-ADTWimFile
#
#-----------------------------------------------------------------------------

function Mount-ADTWimFile
{
    <#
    .SYNOPSIS
        Mounts a WIM file to a specified directory.

    .DESCRIPTION
        Mounts a WIM file to a specified directory. The function supports mounting by image index or image name. It also provides options to forcefully remove existing directories and return the mounted image details.

    .PARAMETER ImagePath
        Path to the WIM file to be mounted.

    .PARAMETER Path
        Directory where the WIM file will be mounted. The directory must be empty and not have a pre-existing WIM mounted.

    .PARAMETER Index
        Index of the image within the WIM file to be mounted.

    .PARAMETER Name
        Name of the image within the WIM file to be mounted.

    .PARAMETER Force
        Forces the removal of the existing directory if it is not empty.

    .PARAMETER PassThru
        If specified, the function will return the results from `Mount-WindowsImage`.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        Microsoft.Dism.Commands.ImageObject

        Returns the mounted image details if the PassThru parameter is specified.

    .EXAMPLE
        Mount-ADTWimFile -ImagePath 'C:\Images\install.wim' -Path 'C:\Mount' -Index 1

        Mounts the first image in the 'install.wim' file to the 'C:\Mount' directory.

    .EXAMPLE
        Mount-ADTWimFile -ImagePath 'C:\Images\install.wim' -Path 'C:\Mount' -Name 'Windows 10 Pro'

        Mounts the image named 'Windows 10 Pro' in the 'install.wim' file to the 'C:\Mount' directory.

    .EXAMPLE
        Mount-ADTWimFile -ImagePath 'C:\Images\install.wim' -Path 'C:\Mount' -Index 1 -Force

        Mounts the first image in the 'install.wim' file to the 'C:\Mount' directory, forcefully removing the existing directory if it is not empty.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Index')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
        [ValidateScript({
                if ($null -eq $_)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ImagePath -ProvidedValue $_ -ExceptionMessage 'The specified input is null.'))
                }
                if (!$_.Exists)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ImagePath -ProvidedValue $_ -ExceptionMessage 'The specified image path cannot be found.'))
                }
                if ([System.Uri]::new($_).IsUnc)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ImagePath -ProvidedValue $_ -ExceptionMessage 'The specified image path cannot be a network share.'))
                }
                return !!$_
            })]
        [System.IO.FileInfo]$ImagePath,

        [Parameter(Mandatory = $true, ParameterSetName = 'Index')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
        [ValidateScript({
                if ($null -eq $_)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified input is null.'))
                }
                if ([System.Uri]::new($_).IsUnc)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified mount path cannot be a network share.'))
                }
                if (& $Script:CommandTable.'Get-ADTMountedWimFile' -Path $_)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified mount path has a pre-existing WIM mounted.'))
                }
                if (& $Script:CommandTable.'Get-ChildItem' -LiteralPath $_ -ErrorAction Ignore)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified mount path is not empty.'))
                }
                return !!$_
            })]
        [System.IO.DirectoryInfo]$Path,

        [Parameter(Mandatory = $true, ParameterSetName = 'Index')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$Index,

        [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false, ParameterSetName = 'Index')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Name')]
        [System.Management.Automation.SwitchParameter]$Force,

        [Parameter(Mandatory = $false, ParameterSetName = 'Index')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Name')]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        # Attempt to get specified WIM image before initialising.
        $null = try
        {
            $PSBoundParameters.Remove('PassThru')
            $PSBoundParameters.Remove('Force')
            $PSBoundParameters.Remove('Path')
            & $Script:CommandTable.'Get-WindowsImage' @PSBoundParameters
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        # Announce commencement.
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Mounting WIM file [$ImagePath] to [$Path]."
        try
        {
            try
            {
                # Provide a warning if this WIM file is already mounted.
                if (($wimFile = & $Script:CommandTable.'Get-ADTMountedWimFile' -ImagePath $ImagePath))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The WIM file [$ImagePath] is already mounted at [$($wimFile.Path)] and will be mounted again." -Severity 2
                }

                # If we're using the force, forcibly remove the existing directory.
                if ([System.IO.Directory]::Exists($Path) -and $Force)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing pre-existing path [$Path] as [-Force] was provided."
                    & $Script:CommandTable.'Remove-Item' -LiteralPath $Path -Force -Confirm:$false
                }

                # If the path doesn't exist, create it.
                if (![System.IO.Directory]::Exists($Path))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating path [$Path] as it does not exist."
                    $Path = [System.IO.Directory]::CreateDirectory($Path).FullName
                }

                # Mount the WIM file.
                $res = & $Script:CommandTable.'Mount-WindowsImage' @PSBoundParameters -Path $Path -ReadOnly -CheckIntegrity
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Successfully mounted WIM file [$ImagePath]."

                # Store the result within the user's ADTSession if there's an active one.
                if (& $Script:CommandTable.'Test-ADTSessionActive')
                {
                    (& $Script:CommandTable.'Get-ADTSession').AddMountedWimFile($ImagePath)
                }

                # Return the result if we're passing through.
                if ($PassThru)
                {
                    return $res
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage 'Error occurred while attemping to mount WIM file.'
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTErrorRecord
#
#-----------------------------------------------------------------------------

function New-ADTErrorRecord
{
    <#
    .SYNOPSIS
        Creates a new ErrorRecord object.

    .DESCRIPTION
        This function creates a new ErrorRecord object with the specified exception, error category, and optional parameters. It allows for detailed error information to be captured and returned to the caller, who can then throw the error.

    .PARAMETER Exception
        The exception object that caused the error.

    .PARAMETER Category
        The category of the error.

    .PARAMETER ErrorId
        The identifier for the error. Default is 'NotSpecified'.

    .PARAMETER TargetObject
        The target object that the error is related to.

    .PARAMETER TargetName
        The name of the target that the error is related to.

    .PARAMETER TargetType
        The type of the target that the error is related to.

    .PARAMETER Activity
        The activity that was being performed when the error occurred.

    .PARAMETER Reason
        The reason for the error.

    .PARAMETER RecommendedAction
        The recommended action to resolve the error.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Management.Automation.ErrorRecord

        This function returns an ErrorRecord object.

    .EXAMPLE
        PS C:\>$exception = [System.Exception]::new("An error occurred.")
        PS C:\>$category = [System.Management.Automation.ErrorCategory]::NotSpecified
        PS C:\>New-ADTErrorRecord -Exception $exception -Category $category -ErrorId "CustomErrorId" -TargetObject $null -TargetName "TargetName" -TargetType "TargetType" -Activity "Activity" -Reason "Reason" -RecommendedAction "RecommendedAction"

        Creates a new ErrorRecord object with the specified parameters.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This function does not change system state.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    [OutputType([System.Management.Automation.ErrorRecord])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Exception]$Exception,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorCategory]$Category,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ErrorId = 'NotSpecified',

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [System.Object]$TargetObject,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TargetName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TargetType,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Activity,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Reason,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$RecommendedAction
    )

    # Instantiate new ErrorRecord object.
    $errRecord = [System.Management.Automation.ErrorRecord]::new($Exception, $ErrorId, $Category, $TargetObject)

    # Add in all optional values, if specified.
    if ($Activity)
    {
        $errRecord.CategoryInfo.Activity = $Activity
    }
    if ($TargetName)
    {
        $errRecord.CategoryInfo.TargetName = $TargetName
    }
    if ($TargetType)
    {
        $errRecord.CategoryInfo.TargetType = $TargetType
    }
    if ($Reason)
    {
        $errRecord.CategoryInfo.Reason = $Reason
    }
    if ($RecommendedAction)
    {
        $errRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($errRecord.Exception.Message)
        $errRecord.ErrorDetails.RecommendedAction = $RecommendedAction
    }

    # Return the ErrorRecord to the caller, who will then throw it.
    return $errRecord
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTFolder
#
#-----------------------------------------------------------------------------

function New-ADTFolder
{
    <#
    .SYNOPSIS
        Create a new folder.

    .DESCRIPTION
        Create a new folder if it does not exist. This function checks if the specified path already exists and creates the folder if it does not. It logs the creation process and handles any errors that may occur during the folder creation.

    .PARAMETER Path
        Path to the new folder to create.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        New-ADTFolder -Path "$env:WinDir\System32"

        Creates a new folder at the specified path if it does not already exist.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        if ([System.IO.Directory]::Exists($Path))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Folder [$Path] already exists."
            return
        }

        try
        {
            try
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating folder [$Path]."
                $null = & $Script:CommandTable.'New-Item' -Path $Path -ItemType Directory -Force
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to create folder [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTMsiTransform
#
#-----------------------------------------------------------------------------

function New-ADTMsiTransform
{
    <#
    .SYNOPSIS
        Create a transform file for an MSI database.

    .DESCRIPTION
        Create a transform file for an MSI database and create/modify properties in the Properties table. This function allows you to specify an existing transform to apply before making changes and to define the path for the new transform file. If the new transform file already exists, it will be deleted before creating a new one.

    .PARAMETER MsiPath
        Specify the path to an MSI file.

    .PARAMETER ApplyTransformPath
        Specify the path to a transform which should be applied to the MSI database before any new properties are created or modified.

    .PARAMETER NewTransformPath
        Specify the path where the new transform file with the desired properties will be created. If a transform file of the same name already exists, it will be deleted before a new one is created.

        Default is:
        a) If -ApplyTransformPath was specified but not -NewTransformPath, then <ApplyTransformPath>.new.mst
        b) If only -MsiPath was specified, then <MsiPath>.mst

    .PARAMETER TransformProperties
        Hashtable which contains calls to Set-ADTMsiProperty for configuring the desired properties which should be included in the new transform file.

        Example hashtable: [Hashtable]$TransformProperties = @{ 'ALLUSERS' = '1' }

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        New-ADTMsiTransform -MsiPath 'C:\Temp\PSADTInstall.msi' -TransformProperties @{
            'ALLUSERS' = '1'
            'AgreeToLicense' = 'Yes'
            'REBOOT' = 'ReallySuppress'
            'RebootYesNo' = 'No'
            'ROOTDRIVE' = 'C:'
        }

        Creates a new transform file for the specified MSI with the given properties.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This function does not change system state.")]
    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!(& $Script:CommandTable.'Test-Path' -Path $_ -PathType Leaf))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName MsiPath -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$MsiPath,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if (!(& $Script:CommandTable.'Test-Path' -Path $_ -PathType Leaf))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ApplyTransformPath -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$ApplyTransformPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$NewTransformPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]$TransformProperties
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Define properties for how the MSI database is opened.
        $msiOpenDatabaseTypes = @{
            OpenDatabaseModeReadOnly = 0
            OpenDatabaseModeTransact = 1
            ViewModifyUpdate = 2
            ViewModifyReplace = 4
            ViewModifyDelete = 6
            TransformErrorNone = 0
            TransformValidationNone = 0
            SuppressApplyTransformErrors = 63
        }
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating a transform file for MSI [$MsiPath]."
        try
        {
            try
            {
                # Create a second copy of the MSI database.
                $MsiParentFolder = & $Script:CommandTable.'Split-Path' -Path $MsiPath -Parent
                $TempMsiPath = & $Script:CommandTable.'Join-Path' -Path $MsiParentFolder -ChildPath ([System.IO.Path]::GetRandomFileName())
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Copying MSI database in path [$MsiPath] to destination [$TempMsiPath]."
                $null = & $Script:CommandTable.'Copy-Item' -LiteralPath $MsiPath -Destination $TempMsiPath -Force

                # Open both copies of the MSI database.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Opening the MSI database [$MsiPath] in read only mode."
                $Installer = & $Script:CommandTable.'New-Object' -ComObject WindowsInstaller.Installer
                $MsiPathDatabase = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Installer -MethodName OpenDatabase -ArgumentList @($MsiPath, $msiOpenDatabaseTypes.OpenDatabaseModeReadOnly)
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Opening the MSI database [$TempMsiPath] in view/modify/update mode."
                $TempMsiPathDatabase = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Installer -MethodName OpenDatabase -ArgumentList @($TempMsiPath, $msiOpenDatabaseTypes.ViewModifyUpdate)

                # If a MSI transform file was specified, then apply it to the temporary copy of the MSI database.
                if ($ApplyTransformPath)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Applying transform file [$ApplyTransformPath] to MSI database [$TempMsiPath]."
                    $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $TempMsiPathDatabase -MethodName ApplyTransform -ArgumentList @($ApplyTransformPath, $msiOpenDatabaseTypes.SuppressApplyTransformErrors)
                }

                # Determine the path for the new transform file that will be generated.
                if (!$NewTransformPath)
                {
                    $NewTransformFileName = if ($ApplyTransformPath)
                    {
                        [System.IO.Path]::GetFileNameWithoutExtension($ApplyTransformPath) + '.new' + [System.IO.Path]::GetExtension($ApplyTransformPath)
                    }
                    else
                    {
                        [System.IO.Path]::GetFileNameWithoutExtension($MsiPath) + '.mst'
                    }
                    $NewTransformPath = & $Script:CommandTable.'Join-Path' -Path $MsiParentFolder -ChildPath $NewTransformFileName
                }

                # Set the MSI properties in the temporary copy of the MSI database.
                foreach ($property in $TransformProperties.GetEnumerator())
                {
                    & $Script:CommandTable.'Set-ADTMsiProperty' -Database $TempMsiPathDatabase -PropertyName $property.Key -PropertyValue $property.Value
                }

                # Commit the new properties to the temporary copy of the MSI database
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $TempMsiPathDatabase -MethodName Commit

                # Reopen the temporary copy of the MSI database in read only mode.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Re-opening the MSI database [$TempMsiPath] in read only mode."
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($TempMsiPathDatabase)
                $TempMsiPathDatabase = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Installer -MethodName OpenDatabase -ArgumentList @($TempMsiPath, $msiOpenDatabaseTypes.OpenDatabaseModeReadOnly)

                # Delete the new transform file path if it already exists.
                if (& $Script:CommandTable.'Test-Path' -LiteralPath $NewTransformPath -PathType Leaf)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "A transform file of the same name already exists. Deleting transform file [$NewTransformPath]."
                    $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $NewTransformPath -Force
                }

                # Generate the new transform file by taking the difference between the temporary copy of the MSI database and the original MSI database.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Generating new transform file [$NewTransformPath]."
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $TempMsiPathDatabase -MethodName GenerateTransform -ArgumentList @($MsiPathDatabase, $NewTransformPath)
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $TempMsiPathDatabase -MethodName CreateTransformSummaryInfo -ArgumentList @($MsiPathDatabase, $NewTransformPath, $msiOpenDatabaseTypes.TransformErrorNone, $msiOpenDatabaseTypes.TransformValidationNone)

                if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $NewTransformPath -PathType Leaf))
                {
                    $naerParams = @{
                        Exception = [System.IO.IOException]::new("Failed to generate transform file in path [$NewTransformPath].")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                        ErrorId = 'MsiTransformFileMissing'
                        TargetObject = $NewTransformPath
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Successfully created new transform file in path [$NewTransformPath]."
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to create new transform file in path [$NewTransformPath]."
        }
        finally
        {
            # Release all COM objects to prevent file locks.
            $null = foreach ($variable in (& $Script:CommandTable.'Get-Variable' -Name TempMsiPathDatabase, MsiPathDatabase, Installer -ValueOnly -ErrorAction Ignore))
            {
                try
                {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($variable)
                }
                catch
                {
                    $null
                }
            }

            # Delete the temporary copy of the MSI database.
            $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $TempMsiPath -Force -ErrorAction Ignore
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTShortcut
#
#-----------------------------------------------------------------------------

function New-ADTShortcut
{
    <#
    .SYNOPSIS
        Creates a new .lnk or .url type shortcut.

    .DESCRIPTION
        Creates a new shortcut .lnk or .url file, with configurable options. This function allows you to specify various parameters such as the target path, arguments, icon location, description, working directory, window style, run as administrator, and hotkey.

    .PARAMETER Path
        Path to save the shortcut.

    .PARAMETER TargetPath
        Target path or URL that the shortcut launches.

    .PARAMETER Arguments
        Arguments to be passed to the target path.

    .PARAMETER IconLocation
        Location of the icon used for the shortcut.

    .PARAMETER IconIndex
        The index of the icon. Executables, DLLs, ICO files with multiple icons need the icon index to be specified. This parameter is an Integer. The first index is 0.

    .PARAMETER Description
        Description of the shortcut.

    .PARAMETER WorkingDirectory
        Working Directory to be used for the target path.

    .PARAMETER WindowStyle
        Windows style of the application. Options: Normal, Maximized, Minimized. Default is: Normal.

    .PARAMETER RunAsAdmin
        Set shortcut to run program as administrator. This option will prompt user to elevate when executing shortcut.

    .PARAMETER Hotkey
        Create a Hotkey to launch the shortcut, e.g. "CTRL+SHIFT+F".

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        New-ADTShortcut -Path "$env:ProgramData\Microsoft\Windows\Start Menu\My Shortcut.lnk" -TargetPath "$env:WinDir\System32\notepad.exe" -IconLocation "$env:WinDir\System32\notepad.exe" -Description 'Notepad' -WorkingDirectory "$env:HomeDrive\$env:HomePath"

        Creates a new shortcut for Notepad with the specified parameters.

    .NOTES
        An active ADT session is NOT required to use this function.

        Url shortcuts only support TargetPath, IconLocation and IconIndex. Other parameters are ignored.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({
                if (![System.IO.Path]::GetExtension($_).ToLower().Equals('.lnk') -and ![System.IO.Path]::GetExtension($_).ToLower().Equals('.url'))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified path does not have the correct extension.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TargetPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Arguments,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IconLocation,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$IconIndex,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Description,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Normal', 'Maximized', 'Minimized')]
        [System.String]$WindowStyle,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$RunAsAdmin,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Hotkey
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        # Make sure .NET's current directory is synced with PowerShell's.
        try
        {
            try
            {
                [System.IO.Directory]::SetCurrentDirectory((& $Script:CommandTable.'Get-Location' -PSProvider FileSystem).ProviderPath)
                $FullPath = [System.IO.Path]::GetFullPath($Path)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Specified path [$Path] is not valid."
            return
        }

        try
        {
            try
            {
                # Make sure directory is present before continuing.
                if (!($PathDirectory = [System.IO.Path]::GetDirectoryName($FullPath)))
                {
                    # The path is root or no filename supplied.
                    if (![System.IO.Path]::GetFileNameWithoutExtension($FullPath))
                    {
                        # No filename supplied.
                        $naerParams = @{
                            Exception = [System.ArgumentException]::new("Specified path [$FullPath] is a directory and not a file.")
                            Category = [System.Management.Automation.ErrorCategory]::InvalidArgument
                            ErrorId = 'ShortcutPathInvalid'
                            TargetObject = $FullPath
                            RecommendedAction = "Please confirm the provided value and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                }
                elseif (!(& $Script:CommandTable.'Test-Path' -LiteralPath $PathDirectory -PathType Container))
                {
                    try
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating shortcut directory [$PathDirectory]."
                        $null = & $Script:CommandTable.'New-Item' -LiteralPath $PathDirectory -ItemType Directory -Force
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to create shortcut directory [$PathDirectory].`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                        throw
                    }
                }

                # Remove any pre-existing shortcut first.
                if (& $Script:CommandTable.'Test-Path' -LiteralPath $FullPath -PathType Leaf)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The shortcut [$FullPath] already exists. Deleting the file..."
                    & $Script:CommandTable.'Remove-ADTFile' -LiteralPath $FullPath
                }

                # Build out the shortcut.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating shortcut [$FullPath]."
                if ($Path -match '\.url$')
                {
                    [String[]]$URLFile = '[InternetShortcut]', "URL=$TargetPath"
                    if ($null -ne $IconIndex)
                    {
                        $URLFile += "IconIndex=$IconIndex"
                    }
                    if ($IconLocation)
                    {
                        $URLFile += "IconFile=$IconLocation"
                    }
                    [System.IO.File]::WriteAllLines($FullPath, $URLFile, [System.Text.UTF8Encoding]::new($false))
                }
                else
                {
                    $shortcut = [System.Activator]::CreateInstance([System.Type]::GetTypeFromProgID('WScript.Shell')).CreateShortcut($FullPath)
                    $shortcut.TargetPath = $TargetPath
                    if ($Arguments)
                    {
                        $shortcut.Arguments = $Arguments
                    }
                    if ($Description)
                    {
                        $shortcut.Description = $Description
                    }
                    if ($WorkingDirectory)
                    {
                        $shortcut.WorkingDirectory = $WorkingDirectory
                    }
                    if ($Hotkey)
                    {
                        $shortcut.Hotkey = $Hotkey
                    }
                    if ($IconLocation)
                    {
                        $shortcut.IconLocation = $IconLocation + ",$IconIndex"
                    }
                    $shortcut.WindowStyle = switch ($WindowStyle)
                    {
                        Normal { 1; break }
                        Maximized { 3; break }
                        Minimized { 7; break }
                        default { 1; break }
                    }

                    # Save the changes.
                    $shortcut.Save()

                    # Set shortcut to run program as administrator.
                    if ($RunAsAdmin)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Setting shortcut to run program as administrator.'
                        $fileBytes = [System.IO.FIle]::ReadAllBytes($FullPath)
                        $fileBytes[21] = $filebytes[21] -bor 32
                        [System.IO.FIle]::WriteAllBytes($FullPath, $fileBytes)
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to create shortcut [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTTemplate
#
#-----------------------------------------------------------------------------

function New-ADTTemplate
{
    <#
    .SYNOPSIS
        Creates a new folder containing a template front end and module folder, ready to customise.

    .DESCRIPTION
        Specify a destination path where a new folder will be created. You also have the option of creating a template for v3 compatibility mode.

    .PARAMETER Destination
        Path where the new folder should be created. Default is the current working directory.

    .PARAMETER Name
        Name of the newly created folder. Default is PSAppDeployToolkit_Version.

    .PARAMETER Version
        Defaults to 4 for the standard v4 template. Use 3 for the v3 compatibility mode template.

    .PARAMETER Show
        Opens the newly created folder in Windows Explorer.

    .PARAMETER Force
        If the destination folder already exists, this switch will force the creation of the new folder.

    .PARAMETER PassThru
        Returns the newly created folder object.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        New-ADTTemplate -Destination 'C:\Temp' -Name 'PSAppDeployToolkitv4'

        Creates a new v4 template named PSAppDeployToolkitv4 under C:\Temp.

    .EXAMPLE
        New-ADTTemplate -Destination 'C:\Temp' -Name 'PSAppDeployToolkitv3' -Version 3

        Creates a new v3 compatibility mode template named PSAppDeployToolkitv3 under C:\Temp.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(SupportsShouldProcess = $false)]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Destination = $ExecutionContext.SessionState.Path.CurrentLocation.Path,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name = "$($MyInvocation.MyCommand.Module.Name)_$($MyInvocation.MyCommand.Module.Version)",

        [Parameter(Mandatory = $false)]
        [ValidateRange(3, 4)]
        [System.Int32]$Version = 4,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Show,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Force,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        # Initialize the function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Resolve the path to handle setups like ".\", etc.
        # We can't use things like a DirectoryInfo cast as .NET doesn't
        # track when the current location in PowerShell has been changed.
        if (($resolvedDest = & $Script:CommandTable.'Resolve-Path' -LiteralPath $Destination -ErrorAction Ignore))
        {
            $Destination = $resolvedDest.Path
        }

        # Set up remaining variables.
        $moduleName = $MyInvocation.MyCommand.Module.Name
        $templatePath = & $Script:CommandTable.'Join-Path' -Path $Destination -ChildPath $Name
        $templateModulePath = if ($Version.Equals(3))
        {
            [System.IO.Path]::Combine($templatePath, 'AppDeployToolkit', $moduleName)
        }
        else
        {
            [System.IO.Path]::Combine($templatePath, $moduleName)
        }
    }

    process
    {
        try
        {
            try
            {
                # If we're running a release module, ensure the psd1 files haven't been tampered with.
                if (($badFiles = & $Script:CommandTable.'Test-ADTReleaseBuildFileValidity' -LiteralPath $Script:PSScriptRoot))
                {
                    $naerParams = @{
                        Exception = [System.InvalidOperationException]::new("One or more files within this module have invalid digital signatures.")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidData
                        ErrorId = 'ADTDataFileSignatureError'
                        TargetObject = $badFiles
                        RecommendedAction = "Please re-download $($MyInvocation.MyCommand.Module.Name) and try again."
                    }
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
                }

                # Create directories.
                if ([System.IO.Directory]::Exists($templatePath) -and [System.IO.Directory]::GetFileSystemEntries($templatePath))
                {
                    if (!$Force)
                    {
                        $naerParams = @{
                            Exception = [System.IO.IOException]::new("Folders [$templatePath] already exists and is not empty.")
                            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                            ErrorId = 'NonEmptySubfolderError'
                            TargetObject = $templatePath
                            RecommendedAction = "Please remove the existing folder, supply a new name, or add the -Force parameter and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                    $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $templatePath -Recurse -Force
                }
                $null = & $Script:CommandTable.'New-Item' -Path "$templatePath\Files" -ItemType Directory -Force
                $null = & $Script:CommandTable.'New-Item' -Path "$templatePath\SupportFiles" -ItemType Directory -Force

                # Add in some empty files to the Files/SupportFiles folders to stop GitHub upload-artifact from dropping the empty folders.
                $null = & $Script:CommandTable.'New-Item' -Name 'Add Setup Files Here.txt' -Path "$templatePath\Files" -ItemType File -Force
                $null = & $Script:CommandTable.'New-Item' -Name 'Add Supporting Files Here.txt' -Path "$templatePath\SupportFiles" -ItemType File -Force

                # Copy in the frontend files and the config/assets/strings.
                & $Script:CommandTable.'Copy-Item' -Path "$Script:PSScriptRoot\Frontend\v$Version\*" -Destination $templatePath -Recurse -Force
                & $Script:CommandTable.'Copy-Item' -LiteralPath "$Script:PSScriptRoot\Assets" -Destination $templatePath -Recurse -Force
                & $Script:CommandTable.'Copy-Item' -LiteralPath "$Script:PSScriptRoot\Config" -Destination $templatePath -Recurse -Force
                & $Script:CommandTable.'Copy-Item' -LiteralPath "$Script:PSScriptRoot\Strings" -Destination $templatePath -Recurse -Force

                # Remove any digital signatures from the ps*1 files.
                & $Script:CommandTable.'Get-ChildItem' -Path "$templatePath\*.ps*1" -Recurse | & {
                    process
                    {
                        if (($sigLine = $(($fileLines = [System.IO.File]::ReadAllLines($_.FullName)) -match '^# SIG # Begin signature block$')))
                        {
                            [System.IO.File]::WriteAllLines($_.FullName, $fileLines[0..($fileLines.IndexOf($sigLine) - 2)])
                        }
                    }
                }

                # Copy in the module files.
                $null = & $Script:CommandTable.'New-Item' -Path $templateModulePath -ItemType Directory -Force
                & $Script:CommandTable.'Copy-Item' -Path "$Script:PSScriptRoot\*" -Destination $templateModulePath -Recurse -Force

                # Make the shipped module and its files read-only.
                $(& $Script:CommandTable.'Get-Item' -LiteralPath $templateModulePath; & $Script:CommandTable.'Get-ChildItem' -LiteralPath $templateModulePath -Recurse) | & {
                    process
                    {
                        $_.Attributes = 'ReadOnly'
                    }
                }

                # Process the generated script to ensure the Import-Module is correct.
                if ($Version.Equals(4))
                {
                    $scriptText = [System.IO.File]::ReadAllText(($scriptFile = "$templatePath\Invoke-AppDeployToolkit.ps1"))
                    $scriptText = $scriptText.Replace("`$PSScriptRoot\..\..\..\$moduleName", "`$PSScriptRoot\$moduleName")
                    [System.IO.File]::WriteAllText($scriptFile, $scriptText, [System.Text.UTF8Encoding]::new($true))
                }

                # Display the newly created folder in Windows Explorer.
                if ($Show)
                {
                    & ([System.IO.Path]::Combine([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Windows), 'explorer.exe')) $templatePath
                }

                # Return a DirectoryInfo object if passing through.
                if ($PassThru)
                {
                    return (& $Script:CommandTable.'Get-Item' -LiteralPath $templatePath)
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTValidateScriptErrorRecord
#
#-----------------------------------------------------------------------------

function New-ADTValidateScriptErrorRecord
{
    <#
    .SYNOPSIS
        Creates a new ErrorRecord for script validation errors.

    .DESCRIPTION
        This function creates a new ErrorRecord object for script validation errors. It takes the parameter name, provided value, exception message, and an optional inner exception to build a detailed error record. This helps in identifying and handling invalid parameter values in scripts.

    .PARAMETER ParameterName
        The name of the parameter that caused the validation error.

    .PARAMETER ProvidedValue
        The value provided for the parameter that caused the validation error.

    .PARAMETER ExceptionMessage
        The message describing the validation error.

    .PARAMETER InnerException
        An optional inner exception that provides more details about the validation error.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Management.Automation.ErrorRecord

        This function returns an ErrorRecord object.

    .EXAMPLE
        PS C:\>$paramName = "FilePath"
        PS C:\>$providedValue = "C:\InvalidPath"
        PS C:\>$exceptionMessage = "The specified path does not exist."
        PS C:\>New-ADTValidateScriptErrorRecord -ParameterName $paramName -ProvidedValue $providedValue -ExceptionMessage $exceptionMessage

        Creates a new ErrorRecord for a validation error with the specified parameters.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = "This function does not change system state.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ParameterName,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [System.Object]$ProvidedValue,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ExceptionMessage,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Exception]$InnerException
    )

    # Build out new ErrorRecord and return it.
    $naerParams = @{
        Exception = if ($InnerException)
        {
            [System.ArgumentException]::new($ExceptionMessage, $ParameterName, $InnerException)
        }
        else
        {
            [System.ArgumentException]::new($ExceptionMessage, $ParameterName)
        }
        Category = [System.Management.Automation.ErrorCategory]::InvalidArgument
        ErrorId = "Invalid$($ParameterName)ParameterValue"
        TargetObject = $ProvidedValue
        TargetName = $ProvidedValue
        TargetType = $(if ($null -ne $ProvidedValue) { $ProvidedValue.GetType().Name })
        RecommendedAction = "Review the supplied $($ParameterName) parameter value and try again."
    }
    return (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
}


#-----------------------------------------------------------------------------
#
# MARK: New-ADTZipFile
#
#-----------------------------------------------------------------------------

function New-ADTZipFile
{
    <#
    .SYNOPSIS
        Create a new zip archive or add content to an existing archive.

    .DESCRIPTION
        Create a new zip archive or add content to an existing archive by using PowerShell's Compress-Archive.

    .PARAMETER Path
        One or more paths to compress. Supports wildcards.

    .PARAMETER LiteralPath
        One or more literal paths to compress.

    .PARAMETER DestinationPath
        The file path for where the zip file should be created.

    .PARAMETER CompressionLevel
        The level of compression to apply to the zip file.

    .PARAMETER Update
        Specifies whether to update an existing zip file or not.

    .PARAMETER Force
        Specifies whether an existing zip file should be overwritten.

    .PARAMETER RemoveSourceAfterArchiving
        Remove the source path after successfully archiving the content.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        New-ADTZipFile -SourceDirectory 'E:\Testing\Logs' -DestinationPath 'E:\Testing\TestingLogs.zip'

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Path,

        [Parameter(Mandatory = $true, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$LiteralPath,

        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [Parameter(Mandatory = $true, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [System.String]$DestinationPath,

        [Parameter(Mandatory = $false, ParameterSetName = 'Path')]
        [Parameter(Mandatory = $false, ParameterSetName = 'LiteralPath')]
        [ValidateSet('Fastest', 'NoCompression', 'Optimal')]
        [System.String]$CompressionLevel,

        [Parameter(Mandatory = $false, ParameterSetName = 'Path')]
        [Parameter(Mandatory = $false, ParameterSetName = 'LiteralPath')]
        [System.Management.Automation.SwitchParameter]$Update,

        [Parameter(Mandatory = $false, ParameterSetName = 'Path')]
        [Parameter(Mandatory = $false, ParameterSetName = 'LiteralPath')]
        [System.Management.Automation.SwitchParameter]$Force,

        [Parameter(Mandatory = $false, ParameterSetName = 'Path')]
        [Parameter(Mandatory = $false, ParameterSetName = 'LiteralPath')]
        [System.Management.Automation.SwitchParameter]$RemoveSourceAfterArchiving
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Remove invalid characters from the supplied filename.
        if (($DestinationArchiveFileName = & $Script:CommandTable.'Remove-ADTInvalidFileNameChars' -Name $DestinationArchiveFileName).Length -eq 0)
        {
            $naerParams = @{
                Exception = [System.ArgumentException]::new('Invalid filename characters replacement resulted into an empty string.', $_)
                Category = [System.Management.Automation.ErrorCategory]::InvalidArgument
                ErrorId = 'DestinationArchiveFileNameInvalid'
                TargetObject = $DestinationArchiveFileName
                RecommendedAction = "Please review the supplied value to '-DestinationArchiveFileName' and try again."
            }
            $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
        }

        # Remove parameters from PSBoundParameters that don't apply to Compress-Archive.
        if ($PSBoundParameters.ContainsKey('RemoveSourceAfterArchiving'))
        {
            $null = $PSBoundParameters.Remove('RemoveSourceAfterArchiving')
        }

        # Get the specified source variable.
        $sourcePath = & $Script:CommandTable.'Get-Variable' -Name $PSCmdlet.ParameterSetName -ValueOnly
    }

    process
    {
        try
        {
            try
            {
                # Get the full destination path where the archive will be stored.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating a zip archive with the requested content at destination path [$DestinationPath]."

                # If the destination archive already exists, delete it if the -OverwriteArchive option was selected.
                if ([System.IO.File]::Exists($DestinationPath) -and $OverwriteArchive)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "An archive at the destination path already exists, deleting file [$DestinationPath]."
                    $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $DestinationPath -Force
                }

                # Create the archive file.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Compressing [$sourcePath] to destination path [$DestinationPath]..."
                & $Script:CommandTable.'Compress-Archive' @PSBoundParameters

                # If option was selected, recursively delete the source directory after successfully archiving the contents.
                if ($RemoveSourceAfterArchiving)
                {
                    try
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Recursively deleting [$sourcePath] as contents have been successfully archived."
                        $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $Directory -Recurse -Force
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to recursively delete [$sourcePath].`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 2
                    }
                }

                # If the archive was created in session 0 or by an Admin, then it may only be readable by elevated users.
                # Apply the parent folder's permissions to the archive file to fix the problem.
                $parentPath = [System.IO.Path]::GetDirectoryName($DestinationPath)
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "If the archive was created in session 0 or by an Admin, then it may only be readable by elevated users. Apply permissions from parent folder [$parentPath] to file [$DestinationPath]."
                try
                {
                    & $Script:CommandTable.'Set-Acl' -LiteralPath $DestinationPath -AclObject (& $Script:CommandTable.'Get-Acl' -Path $parentPath)
                }
                catch
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to apply parent folder's [$parentPath] permissions to file [$DestinationPath].`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 2
                }
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to archive the requested file(s)."
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Open-ADTSession
#
#-----------------------------------------------------------------------------

function Open-ADTSession
{
    <#
    .SYNOPSIS
        Opens a new ADT session.

    .DESCRIPTION
        This function initializes and opens a new ADT session with the specified parameters. It handles the setup of the session environment and processes any callbacks defined for the session. If the session fails to open, it handles the error and closes the session if necessary.

    .PARAMETER SessionState
        Caller's SessionState.

    .PARAMETER DeploymentType
        Specifies the type of deployment: Install, Uninstall, or Repair.

    .PARAMETER DeployMode
        Specifies the deployment mode: Interactive, NonInteractive, or Silent.

    .PARAMETER AllowRebootPassThru
        Allows reboot pass-through.

    .PARAMETER TerminalServerMode
        Enables Terminal Server mode.

    .PARAMETER DisableLogging
        Disables logging for the session.

    .PARAMETER AppVendor
        Specifies the application vendor.

    .PARAMETER AppName
        Specifies the application name.

    .PARAMETER AppVersion
        Specifies the application version.

    .PARAMETER AppArch
        Specifies the application architecture.

    .PARAMETER AppLang
        Specifies the application language.

    .PARAMETER AppRevision
        Specifies the application revision.

    .PARAMETER AppSuccessExitCodes
        Specifies the application exit codes.

    .PARAMETER AppRebootExitCodes
        Specifies the application reboot codes.

    .PARAMETER AppScriptVersion
        Specifies the application script version.

    .PARAMETER AppScriptDate
        Specifies the application script date.

    .PARAMETER AppScriptAuthor
        Specifies the application script author.

    .PARAMETER InstallName
        Specifies the install name.

    .PARAMETER InstallTitle
        Specifies the install title.

    .PARAMETER DeployAppScriptFriendlyName
        Specifies the friendly name of the deploy application script.

    .PARAMETER DeployAppScriptVersion
        Specifies the version of the deploy application script.

    .PARAMETER DeployAppScriptDate
        Specifies the date of the deploy application script.

    .PARAMETER DeployAppScriptParameters
        Specifies the parameters for the deploy application script.

    .PARAMETER ScriptDirectory
        Specifies the base path for Files and SupportFiles.

    .PARAMETER DirFiles
        Specifies the override path to Files.

    .PARAMETER DirSupportFiles
        Specifies the override path to SupportFiles.

    .PARAMETER DefaultMsiFile
        Specifies the default MSI file.

    .PARAMETER DefaultMstFile
        Specifies the default MST file.

    .PARAMETER DefaultMspFiles
        Specifies the default MSP files.

    .PARAMETER DisableDefaultMsiProcessList
        Specifies that the zero-config MSI code should not gather process names from the MSI file.

    .PARAMETER LogName
        Specifies an override for the default-generated log file name.

    .PARAMETER SessionClass
        Specifies an override for PSADT.Module.DeploymentSession class. Use this if you're deriving a class inheriting off PSAppDeployToolkit's base.

    .PARAMETER ForceWimDetection
        Specifies that WIM files should be detected and mounted during session initialization, irrespective of whether any App values are provided.

    .PARAMETER PassThru
        Passes the session object through the pipeline.

    .PARAMETER UnboundArguments
        Captures any additional arguments passed to the function.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        ADTSession

        This function returns the session object if -PassThru is specified.

    .EXAMPLE
        Open-ADTSession -SessionState $ExecutionContext.SessionState -DeploymentType "Install" -DeployMode "Interactive"

        Opens a new ADT session with the specified parameters.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SessionState]$SessionState,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Parameter')]
        [ValidateNotNullOrEmpty()]
        [PSADT.Module.DeploymentType]$DeploymentType,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Parameter')]
        [ValidateNotNullOrEmpty()]
        [PSADT.Module.DeployMode]$DeployMode,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Parameter')]
        [System.Management.Automation.SwitchParameter]$AllowRebootPassThru,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Parameter')]
        [System.Management.Automation.SwitchParameter]$TerminalServerMode,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Parameter')]
        [System.Management.Automation.SwitchParameter]$DisableLogging,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$AppVendor,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$AppName,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$AppVersion,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$AppArch,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$AppLang,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$AppRevision,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [ValidateNotNullOrEmpty()]
        [System.Version]$AppScriptVersion,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [ValidateNotNullOrEmpty()]
        [System.DateTime]$AppScriptDate,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [ValidateNotNullOrEmpty()]
        [System.String]$AppScriptAuthor,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$InstallName,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyString()]
        [System.String]$InstallTitle,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [ValidateNotNullOrEmpty()]
        [System.String]$DeployAppScriptFriendlyName,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [ValidateNotNullOrEmpty()]
        [System.Version]$DeployAppScriptVersion,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [ValidateNotNullOrEmpty()]
        [System.DateTime]$DeployAppScriptDate,

        [Parameter(Mandatory = $false, HelpMessage = 'Frontend Variable')]
        [AllowEmptyCollection()]
        [System.Collections.Generic.Dictionary[System.String, System.Object]]$DeployAppScriptParameters,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]$AppSuccessExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]$AppRebootExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ScriptDirectory -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName ScriptDirectory -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String[]]$ScriptDirectory,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName DirFiles -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName DirFiles -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String]$DirFiles,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName DirSupportFiles -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (![System.IO.Directory]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName DirSupportFiles -ProvidedValue $_ -ExceptionMessage 'The specified directory does not exist.'))
                }
                return $_
            })]
        [System.String]$DirSupportFiles,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DefaultMsiFile,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$DefaultMstFile,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$DefaultMspFiles,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DisableDefaultMsiProcessList,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if ([System.String]::IsNullOrWhiteSpace($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName LogName -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if ([System.IO.Path]::GetExtension($_) -ne '.log')
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName LogName -ProvidedValue $_ -ExceptionMessage 'The specified name does not have a [.log] extension.'))
                }
                return $_
            })]
        [System.String]$LogName,

        [Parameter(Mandatory = $false, DontShow = $true)]
        [ValidateScript({
                if ($null -eq $_)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName SessionClass -ProvidedValue $_ -ExceptionMessage 'The specified input is null or empty.'))
                }
                if (!$_.BaseType.Equals([PSADT.Module.DeploymentSession]))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName SessionClass -ProvidedValue $_ -ExceptionMessage 'The specified type is not derived from the DeploymentSession base class.'))
                }
                return $_
            })]
        [System.Type]$SessionClass = [PSADT.Module.DeploymentSession],

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ForceWimDetection,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false, ValueFromRemainingArguments = $true, DontShow = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Generic.List[System.Object]]$UnboundArguments
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $adtSession = $null
        $errRecord = $null

        # Determine whether this session is to be in compatibility mode.
        $compatibilityMode = & $Script:CommandTable.'Test-ADTNonNativeCaller'
        $callerInvocation = (& $Script:CommandTable.'Get-PSCallStack')[1].InvocationInfo
        $noExitOnClose = !$callerInvocation.MyCommand.CommandType.Equals([System.Management.Automation.CommandTypes]::ExternalScript) -and !([System.Environment]::GetCommandLineArgs() -eq '-NonInteractive')

        # Set up the ScriptDirectory if one wasn't provided.
        if (!$PSBoundParameters.ContainsKey('ScriptDirectory'))
        {
            [System.String[]]$PSBoundParameters.ScriptDirectory = if (![System.String]::IsNullOrWhiteSpace(($scriptRoot = $SessionState.PSVariable.GetValue('PSScriptRoot', $null))))
            {
                if ($compatibilityMode)
                {
                    [System.IO.Directory]::GetParent($scriptRoot).FullName
                }
                else
                {
                    $scriptRoot
                }
            }
            else
            {
                $ExecutionContext.SessionState.Path.CurrentLocation.Path
            }
        }
    }

    process
    {
        # If this function is being called from the console or by AppDeployToolkitMain.ps1, clear all previous sessions and go for full re-initialization.
        if (([System.String]::IsNullOrWhiteSpace($callerInvocation.InvocationName) -and [System.String]::IsNullOrWhiteSpace($callerInvocation.Line)) -or $compatibilityMode)
        {
            $Script:ADT.Sessions.Clear()
            $Script:ADT.Initialized = $false
        }
        $firstSession = !$Script:ADT.Sessions.Count

        # Commence the opening process.
        try
        {
            try
            {
                # Initialize the module before opening the first session.
                if ($firstSession -and !$Script:ADT.Initialized)
                {
                    & $Script:CommandTable.'Initialize-ADTModule' -ScriptDirectory $PSBoundParameters.ScriptDirectory
                }

                # Instantiate the new session. The constructor will handle adding the session to the module's list.
                $adtSession = $SessionClass::new($PSBoundParameters, $noExitOnClose, $(if ($compatibilityMode) { $SessionState }))

                # Invoke all callbacks.
                foreach ($callback in $(if ($firstSession) { $Script:ADT.Callbacks.Starting }; $Script:ADT.Callbacks.Opening))
                {
                    & $callback
                }

                # Add any unbound arguments into the $adtSession object as PSNoteProperty objects.
                if ($PSBoundParameters.ContainsKey('UnboundArguments'))
                {
                    (& $Script:CommandTable.'Convert-ADTValuesFromRemainingArguments' -RemainingArguments $UnboundArguments).GetEnumerator() | & {
                        begin
                        {
                            $adtSessionProps = $adtSession.PSObject.Properties
                        }

                        process
                        {
                            $adtSessionProps.Add([System.Management.Automation.PSNoteProperty]::new($_.Key, $_.Value))
                        }
                    }
                }

                # Export the environment table to variables within the caller's scope.
                if ($firstSession)
                {
                    & $Script:CommandTable.'Export-ADTEnvironmentTableToSessionState' -SessionState $SessionState
                }

                # Change the install phase since we've finished initialising. This should get overwritten shortly.
                $adtSession.InstallPhase = 'Execution'
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord ($errRecord = $_) -LogMessage "Failure occurred while opening new deployment session."
        }
        finally
        {
            # Terminate early if we have an active session that failed to open properly.
            if ($errRecord)
            {
                if (!$adtSession)
                {
                    & $Script:CommandTable.'Exit-ADTInvocation' -ExitCode $(if (!$noExitOnClose) { 60008 })
                }
                else
                {
                    & $Script:CommandTable.'Close-ADTSession' -ExitCode 60008
                }
            }
        }

        # Return the most recent session if passing through.
        if ($PassThru)
        {
            return $adtSession
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Out-ADTPowerShellEncodedCommand
#
#-----------------------------------------------------------------------------

function Out-ADTPowerShellEncodedCommand
{
    <#
    .SYNOPSIS
        Encodes a PowerShell command into a Base64 string.

    .DESCRIPTION
        This function takes a PowerShell command as input and encodes it into a Base64 string. This is useful for passing commands to PowerShell through mechanisms that require encoded input.

    .PARAMETER Command
        The PowerShell command to be encoded.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        This function returns the encoded Base64 string representation of the input command.

    .EXAMPLE
        Out-ADTPowerShellEncodedCommand -Command 'Get-Process'

        Encodes the "Get-Process" command into a Base64 string.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Command
    )

    return [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($Command))
}


#-----------------------------------------------------------------------------
#
# MARK: Register-ADTDll
#
#-----------------------------------------------------------------------------

function Register-ADTDll
{
    <#
    .SYNOPSIS
        Register a DLL file.

    .DESCRIPTION
        This function registers a DLL file using regsvr32.exe. It ensures that the specified DLL file exists before attempting to register it. If the file does not exist, it throws an error.

    .PARAMETER FilePath
        Path to the DLL file.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return objects.

    .EXAMPLE
        Register-ADTDll -FilePath "C:\Test\DcTLSFileToDMSComp.dll"

        Registers the specified DLL file.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            & $Script:CommandTable.'Invoke-ADTRegSvr32' @PSBoundParameters -Action Register
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTContentFromCache
#
#-----------------------------------------------------------------------------

function Remove-ADTContentFromCache
{
    <#
    .SYNOPSIS
        Removes the toolkit content from the cache folder on the local machine and reverts the $adtSession.DirFiles and $adtSession.SupportFiles directory.

    .DESCRIPTION
        This function removes the toolkit content from the cache folder on the local machine. It also reverts the $adtSession.DirFiles and $adtSession.SupportFiles directory to their original state. If the specified cache folder does not exist, it logs a message and exits.

    .PARAMETER Path
        The path to the software cache folder.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return objects.

    .EXAMPLE
        Remove-ADTContentFromCache -Path "$envWinDir\Temp\PSAppDeployToolkit"

        Removes the toolkit content from the specified cache folder.

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Path = "$((& $Script:CommandTable.'Get-ADTConfig').Toolkit.CachePath)\$((& $Script:CommandTable.'Get-ADTSession').installName)"
    )

    begin
    {
        try
        {
            $adtSession = & $Script:CommandTable.'Get-ADTSession'
            $parentPath = $adtSession.ScriptDirectory
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        if (![System.IO.Directory]::Exists($Path))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Cache folder [$Path] does not exist."
            return
        }

        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing cache folder [$Path]."
        try
        {
            try
            {
                & $Script:CommandTable.'Remove-Item' -Path $Path -Recurse
                $adtSession.DirFiles = (& $Script:CommandTable.'Join-Path' -Path $parentPath -ChildPath Files)
                $adtSession.DirSupportFiles = (& $Script:CommandTable.'Join-Path' -Path $parentPath -ChildPath SupportFiles)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to remove cache folder [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTEdgeExtension
#
#-----------------------------------------------------------------------------

function Remove-ADTEdgeExtension
{
    <#
    .SYNOPSIS
        Removes an extension for Microsoft Edge using the ExtensionSettings policy.

    .DESCRIPTION
        This function removes an extension for Microsoft Edge using the ExtensionSettings policy: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-manage-extensions-ref-guide.

        This enables Edge Extensions to be installed and managed like applications, enabling extensions to be pushed to specific devices or users alongside existing GPO/Intune extension policies.

        This should not be used in conjunction with Edge Management Service which leverages the same registry key to configure Edge extensions.

    .PARAMETER ExtensionID
        The ID of the extension to remove.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return objects.

    .EXAMPLE
        Remove-ADTEdgeExtension -ExtensionID "extensionID"

        Removes the specified extension from Microsoft Edge.

    .NOTES
        An active ADT session is NOT required to use this function.

        This function is provided as a template to remove an extension for Microsoft Edge. This should not be used in conjunction with Edge Management Service which leverages the same registry key to configure Edge extensions.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ExtensionID
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing extension with ID [$ExtensionID]."
        try
        {
            try
            {
                # Return early if the extension isn't installed.
                if (!($installedExtensions = & $Script:CommandTable.'Get-ADTEdgeExtensions').PSObject.Properties -or ($installedExtensions.PSObject.Properties.Name -notcontains $ExtensionID))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Extension with ID [$ExtensionID] is not configured. Removal not required."
                    return
                }

                # If the deploymentmode is Remove, remove the extension from the list.
                $installedExtensions.PSObject.Properties.Remove($ExtensionID)
                $null = & $Script:CommandTable.'Set-ADTRegistryKey' -Key Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge -Name ExtensionSettings -Value ($installedExtensions | & $Script:CommandTable.'ConvertTo-Json' -Compress)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTFile
#
#-----------------------------------------------------------------------------

function Remove-ADTFile
{
    <#
    .SYNOPSIS
        Removes one or more items from a given path on the filesystem.

    .DESCRIPTION
        This function removes one or more items from a given path on the filesystem. It can handle both wildcard paths and literal paths. If the specified path does not exist, it logs a warning instead of throwing an error. The function can also delete items recursively if the Recurse parameter is specified.

    .PARAMETER Path
        Specifies the path on the filesystem to be resolved. The value of Path will accept wildcards. Will accept an array of values.

    .PARAMETER LiteralPath
        Specifies the path on the filesystem to be resolved. The value of LiteralPath is used exactly as it is typed; no characters are interpreted as wildcards. Will accept an array of values.

    .PARAMETER Recurse
        Deletes the files in the specified location(s) and in all child items of the location(s).

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTFile -Path 'C:\Windows\Downloaded Program Files\Temp.inf'

        Removes the specified file.

    .EXAMPLE
        Remove-ADTFile -LiteralPath 'C:\Windows\Downloaded Program Files' -Recurse

        Removes the specified folder and all its contents recursively.

    .NOTES
        An active ADT session is NOT required to use this function.

        This function continues on received errors by default. To have the function stop on an error, please provide `-ErrorAction Stop` on the end of your call.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'LiteralPath', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Path', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Path,

        [Parameter(Mandatory = $true, ParameterSetName = 'LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$LiteralPath,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        foreach ($Item in $PSBoundParameters[$PSCmdlet.ParameterSetName])
        {
            # Resolve the specified path, if the path does not exist, display a warning instead of an error.
            try
            {
                try
                {
                    $Item = if ($PSCmdlet.ParameterSetName -eq 'Path')
                    {
                        (& $Script:CommandTable.'Resolve-Path' -Path $Item).Path
                    }
                    else
                    {
                        (& $Script:CommandTable.'Resolve-Path' -LiteralPath $Item).Path
                    }
                }
                catch [System.Management.Automation.ItemNotFoundException]
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to resolve the path [$Item] because it does not exist." -Severity 2
                    continue
                }
                catch [System.Management.Automation.DriveNotFoundException]
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to resolve the path [$Item] because the drive does not exist." -Severity 2
                    continue
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to resolve the path for deletion [$Item]."
                continue
            }

            # Delete specified path if it was successfully resolved.
            try
            {
                try
                {
                    if (& $Script:CommandTable.'Test-Path' -LiteralPath $Item -PathType Container)
                    {
                        if (!$Recurse)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Skipping folder [$Item] because the Recurse switch was not specified."
                            continue
                        }
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting file(s) recursively in path [$Item]..."
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting file in path [$Item]..."
                    }
                    $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $Item -Recurse:$Recurse -Force
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to delete items in path [$Item]."
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTFileFromUserProfiles
#
#-----------------------------------------------------------------------------

function Remove-ADTFileFromUserProfiles
{
    <#
    .SYNOPSIS
        Removes one or more items from each user profile on the system.

    .DESCRIPTION
        This function removes one or more items from each user profile on the system. It can handle both wildcard paths and literal paths. If the specified path does not exist, it logs a warning instead of throwing an error. The function can also delete items recursively if the Recurse parameter is specified. Additionally, it allows excluding specific NT accounts, system profiles, service profiles, and the default user profile.

    .PARAMETER Path
        Specifies the path to append to the root of the user profile to be resolved. The value of Path will accept wildcards. Will accept an array of values.

    .PARAMETER LiteralPath
        Specifies the path to append to the root of the user profile to be resolved. The value of LiteralPath is used exactly as it is typed; no characters are interpreted as wildcards. Will accept an array of values.

    .PARAMETER Recurse
        Deletes the files in the specified location(s) and in all child items of the location(s).

    .PARAMETER ExcludeNTAccount
        Specify NT account names in Domain\Username format to exclude from the list of user profiles.

    .PARAMETER ExcludeDefaultUser
        Exclude the Default User. Default is: $false.

    .PARAMETER IncludeSystemProfiles
        Include system profiles: SYSTEM, LOCAL SERVICE, NETWORK SERVICE. Default is: $false.

    .PARAMETER IncludeServiceProfiles
        Include service profiles where NTAccount begins with NT SERVICE. Default is: $false.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTFileFromUserProfiles -Path "AppData\Roaming\MyApp\config.txt"

        Removes the specified file from each user profile on the system.

    .EXAMPLE
        Remove-ADTFileFromUserProfiles -Path "AppData\Local\MyApp" -Recurse

        Removes the specified folder and all its contents recursively from each user profile on the system.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'LiteralPath', Justification = "This parameter is accessed programmatically via the ParameterSet it's within, which PSScriptAnalyzer doesn't understand.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Path', Justification = "This parameter is accessed programmatically via the ParameterSet it's within, which PSScriptAnalyzer doesn't understand.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
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
        [System.Management.Automation.SwitchParameter]$ExcludeDefaultUser,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeSystemProfiles,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$IncludeServiceProfiles
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $RemoveFileSplat = @{
            Recurse = $Recurse
        }
        $GetUserProfileSplat = @{
            IncludeSystemProfiles = $IncludeSystemProfiles
            IncludeServiceProfiles = $IncludeServiceProfiles
            ExcludeDefaultUser = $ExcludeDefaultUser
        }
        if ($ExcludeNTAccount)
        {
            $GetUserProfileSplat.ExcludeNTAccount = $ExcludeNTAccount
        }

        # Store variable based on ParameterSetName.
        $pathVar = & $Script:CommandTable.'Get-Variable' -Name $PSCmdlet.ParameterSetName
    }

    process
    {
        foreach ($UserProfilePath in (& $Script:CommandTable.'Get-ADTUserProfiles' @GetUserProfileSplat).ProfilePath)
        {
            $RemoveFileSplat.Path = $pathVar.Value | & { process { [System.IO.Path]::Combine($UserProfilePath, $_) } }
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing $($pathVar.Name) [$($pathVar.Value)] from $UserProfilePath`:"
            try
            {
                try
                {
                    & $Script:CommandTable.'Remove-ADTFile' @RemoveFileSplat
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTFolder
#
#-----------------------------------------------------------------------------

function Remove-ADTFolder
{
    <#
    .SYNOPSIS
        Remove folder and files if they exist.

    .DESCRIPTION
        This function removes a folder and all files within it, with or without recursion, in a given path. If the specified folder does not exist, it logs a warning instead of throwing an error. The function can also delete items recursively if the DisableRecursion parameter is not specified.

    .PARAMETER Path
        Path to the folder to remove.

    .PARAMETER DisableRecursion
        Disables recursion while deleting.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTFolder -Path "$envWinDir\Downloaded Program Files"

        Deletes all files and subfolders in the Windows\Downloads Program Files folder.

    .EXAMPLE
        Remove-ADTFolder -Path "$envTemp\MyAppCache" -DisableRecursion

        Deletes all files in the Temp\MyAppCache folder but does not delete any subfolders.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.DirectoryInfo]$Path,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DisableRecursion
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        # Return early if the folder doesn't exist.
        if (!($Path | & $Script:CommandTable.'Test-Path' -PathType Container))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Folder [$Path] does not exist."
            return
        }

        try
        {
            try
            {
                # With -Recurse, we can just send it and return early.
                if (!$DisableRecursion)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting folder [$Path] recursively..."
                    & $Script:CommandTable.'Invoke-ADTCommandWithRetries' -Command $Script:CommandTable.'Remove-Item' -LiteralPath $Path -Force -Recurse
                    return
                }

                # Without recursion, we can only send it if the folder has no items as Remove-Item will ask for confirmation without recursion.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting folder [$Path] without recursion..."
                if (!($ListOfChildItems = & $Script:CommandTable.'Get-ChildItem' -LiteralPath $Path -Force))
                {
                    & $Script:CommandTable.'Invoke-ADTCommandWithRetries' -Command $Script:CommandTable.'Remove-Item' -LiteralPath $Path -Force
                    return
                }

                # We must have some subfolders, let's see what we can do.
                $SubfoldersSkipped = foreach ($item in $ListOfChildItems)
                {
                    # Check whether this item is a folder
                    if ($item -is [System.IO.DirectoryInfo])
                    {
                        # Item is a folder. Check if its empty.
                        if (($item | & $Script:CommandTable.'Get-ChildItem' -Force | & $Script:CommandTable.'Measure-Object').Count -eq 0)
                        {
                            # The folder is empty, delete it
                            & $Script:CommandTable.'Invoke-ADTCommandWithRetries' -Command $Script:CommandTable.'Remove-Item' -LiteralPath $item.FullName -Force
                        }
                        else
                        {
                            # Folder is not empty, skip it.
                            $item
                        }
                    }
                    else
                    {
                        # Item is a file. Delete it.
                        & $Script:CommandTable.'Invoke-ADTCommandWithRetries' -Command $Script:CommandTable.'Remove-Item' -LiteralPath $item.FullName -Force
                    }
                }
                if ($SubfoldersSkipped)
                {
                    $naerParams = @{
                        Exception = [System.IO.IOException]::new("The following folders are not empty ['$($SubfoldersSkipped.FullName.Replace($Path.FullName, $null) -join "'; '")'].")
                        Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                        ErrorId = 'NonEmptySubfolderError'
                        TargetObject = $SubfoldersSkipped
                        RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to delete folder(s) and file(s) from path [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTInvalidFileNameChars
#
#-----------------------------------------------------------------------------

function Remove-ADTInvalidFileNameChars
{
    <#
    .SYNOPSIS
        Remove invalid characters from the supplied string.

    .DESCRIPTION
        This function removes invalid characters from the supplied string and returns a valid filename as a string. It ensures that the resulting string does not contain any characters that are not allowed in filenames. This function should not be used for entire paths as '\' is not a valid filename character.

    .PARAMETER Name
        Text to remove invalid filename characters from.

    .INPUTS
        System.String

        A string containing invalid filename characters.

    .OUTPUTS
        System.String

        Returns the input string with the invalid characters removed.

    .EXAMPLE
        Remove-ADTInvalidFileNameChars -Name "Filename/\1"

        Removes invalid filename characters from the string "Filename/\1".

    .NOTES
        An active ADT session is NOT required to use this function.

        This function always returns a string; however, it can be empty if the name only contains invalid characters. Do not use this command for an entire path as '\' is not a valid filename character.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [System.String]$Name
    )

    process
    {
        return ($Name.Trim() -replace "[$([System.Text.RegularExpressions.Regex]::Escape([System.String]::Join($null, [System.IO.Path]::GetInvalidFileNameChars())))]")
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTRegistryKey
#
#-----------------------------------------------------------------------------

function Remove-ADTRegistryKey
{
    <#
    .SYNOPSIS
        Deletes the specified registry key or value.

    .DESCRIPTION
        This function deletes the specified registry key or value. It can handle both registry keys and values, and it supports recursive deletion of registry keys. If the SID parameter is specified, it converts HKEY_CURRENT_USER registry keys to the HKEY_USERS\$SID format, allowing for the manipulation of HKCU registry settings for all users on the system.

    .PARAMETER Key
        Path of the registry key to delete.

    .PARAMETER Name
        Name of the registry value to delete.

    .PARAMETER Recurse
        Delete registry key recursively.

    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.

        Specify this parameter from the Invoke-ADTAllUsersRegistryAction function to read/edit HKCU registry settings for all users on the system.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTRegistryKey -Key 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce'

        Deletes the specified registry key.

    .EXAMPLE
        Remove-ADTRegistryKey -Key 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'RunAppInstall'

        Deletes the specified registry value.

    .EXAMPLE
        Remove-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Example' -Name '(Default)'

        Deletes the default registry value in the specified key.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Recurse,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        try
        {
            try
            {
                # If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID.
                $Key = if ($PSBoundParameters.ContainsKey('SID'))
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -SID $SID
                }
                else
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key
                }

                if (!$Name)
                {
                    if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $Key))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to delete registry key [$Key] because it does not exist." -Severity 2
                        return
                    }

                    if ($Recurse)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting registry key recursively [$Key]."
                        $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $Key -Force -Recurse
                    }
                    elseif (!(& $Script:CommandTable.'Get-ChildItem' -LiteralPath $Key))
                    {
                        # Check if there are subkeys of $Key, if so, executing Remove-Item will hang. Avoiding this with Get-ChildItem.
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting registry key [$Key]."
                        $null = & $Script:CommandTable.'Remove-Item' -LiteralPath $Key -Force
                    }
                    else
                    {
                        $naerParams = @{
                            Exception = [System.InvalidOperationException]::new("Unable to delete child key(s) of [$Key] without [-Recurse] switch.")
                            Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                            ErrorId = 'SubKeyRecursionError'
                            TargetObject = $Key
                            RecommendedAction = "Please run this command again with [-Recurse]."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                }
                else
                {
                    if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $Key))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to delete registry value [$Key] [$Name] because registry key does not exist." -Severity 2
                        return
                    }
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Deleting registry value [$Key] [$Name]."
                    if ($Name -eq '(Default)')
                    {
                        # Remove (Default) registry key value with the following workaround because Remove-ItemProperty cannot remove the (Default) registry key value.
                        $null = (& $Script:CommandTable.'Get-Item' -LiteralPath $Key).OpenSubKey('', 'ReadWriteSubTree').DeleteValue('')
                    }
                    else
                    {
                        $null = & $Script:CommandTable.'Remove-ItemProperty' -LiteralPath $Key -Name $Name -Force
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch [System.Management.Automation.PSArgumentException]
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to delete registry value [$Key] [$Name] because it does not exist." -Severity 2
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to delete registry $(("key [$Key]", "value [$Key] [$Name]")[!!$Name])."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTSessionClosingCallback
#
#-----------------------------------------------------------------------------

function Remove-ADTSessionClosingCallback
{
    <#
    .SYNOPSIS
        Removes a callback function from the ADT session closing event.

    .DESCRIPTION
        This function removes a specified callback function from the ADT session closing event. The callback function must be provided as a parameter. If the operation fails, it throws a terminating error.

    .PARAMETER Callback
        The callback function to remove from the ADT session closing event.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTSessionClosingCallback -Callback (Get-Command -Name 'MyCallbackFunction')

        Removes the specified callback function from the ADT session closing event.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Closing -Action Remove @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTSessionFinishingCallback
#
#-----------------------------------------------------------------------------

function Remove-ADTSessionFinishingCallback
{
    <#
    .SYNOPSIS
        Removes a callback function from the ADT session finishing event.

    .DESCRIPTION
        This function removes a specified callback function from the ADT session finishing event. The callback function must be provided as a parameter. If the operation fails, it throws a terminating error.

    .PARAMETER Callback
        The callback function to remove from the ADT session finishing event.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTSessionFinishingCallback -Callback (Get-Command -Name 'MyCallbackFunction')

        Removes the specified callback function from the ADT session finishing event.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Finishing -Action Remove @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTSessionOpeningCallback
#
#-----------------------------------------------------------------------------

function Remove-ADTSessionOpeningCallback
{
    <#
    .SYNOPSIS
        Removes a callback function from the ADT session opening event.

    .DESCRIPTION
        This function removes a specified callback function from the ADT session opening event. The callback function must be provided as a parameter. If the operation fails, it throws a terminating error.

    .PARAMETER Callback
        The callback function to remove from the ADT session opening event.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTSessionOpeningCallback -Callback (Get-Command -Name 'MyCallbackFunction')

        Removes the specified callback function from the ADT session opening event.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Opening -Action Remove @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Remove-ADTSessionStartingCallback
#
#-----------------------------------------------------------------------------

function Remove-ADTSessionStartingCallback
{
    <#
    .SYNOPSIS
        Removes a callback function from the ADT session starting event.

    .DESCRIPTION
        This function removes a specified callback function from the ADT session starting event. The callback function must be provided as a parameter. If the operation fails, it throws a terminating error.

    .PARAMETER Callback
        The callback function to remove from the ADT session starting event.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Remove-ADTSessionStartingCallback -Callback (Get-Command -Name 'MyCallbackFunction')

        Removes the specified callback function from the ADT session starting event.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.CommandInfo[]]$Callback
    )

    # Send it off to the backend function.
    try
    {
        & $Script:CommandTable.'Invoke-ADTSessionCallbackOperation' -Type Starting -Action Remove @PSBoundParameters
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Reset-ADTDeferHistory
#
#-----------------------------------------------------------------------------

function Reset-ADTDeferHistory
{
    <#
    .SYNOPSIS
        Reset the history of deferrals in the registry for the current application.

    .DESCRIPTION
        Reset the history of deferrals in the registry for the current application.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Reset-DeferHistory

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com

    #>

    [CmdletBinding()]
    param
    (
    )

    try
    {
        (& $Script:CommandTable.'Get-ADTSession').ResetDeferHistory()
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Resolve-ADTErrorRecord
#
#-----------------------------------------------------------------------------

function Resolve-ADTErrorRecord
{
    <#
    .SYNOPSIS
        Enumerates ErrorRecord details.

    .DESCRIPTION
        Enumerates an ErrorRecord, or a collection of ErrorRecord properties. This function can filter and display specific properties of the ErrorRecord, and can exclude certain parts of the error details.

    .PARAMETER ErrorRecord
        The ErrorRecord to resolve. For usage in a catch block, you'd use the automatic variable `$PSItem`. For usage out of a catch block, you can access the global $Error array's first error (on index 0).

    .PARAMETER Property
        The list of properties to display from the ErrorRecord. Use "*" to display all properties.

        Default list of error properties is: Message, FullyQualifiedErrorId, ScriptStackTrace, PositionMessage, InnerException

    .PARAMETER ExcludeErrorRecord
        Exclude ErrorRecord details as represented by $ErrorRecord.

    .PARAMETER ExcludeErrorInvocation
        Exclude ErrorRecord invocation information as represented by $ErrorRecord.InvocationInfo.

    .PARAMETER ExcludeErrorException
        Exclude ErrorRecord exception details as represented by $ErrorRecord.Exception.

    .PARAMETER ExcludeErrorInnerException
        Exclude ErrorRecord inner exception details as represented by $ErrorRecord.Exception.InnerException. Will retrieve all inner exceptions if there is more than one.

    .INPUTS
        System.Management.Automation.ErrorRecord

        Accepts one or more ErrorRecord objects via the pipeline.

    .OUTPUTS
        System.String

        Displays the ErrorRecord details.

    .EXAMPLE
        Resolve-ADTErrorRecord

        Enumerates the details of the last ErrorRecord.

    .EXAMPLE
        Resolve-ADTErrorRecord -Property *

        Enumerates all properties of the last ErrorRecord.

    .EXAMPLE
        Resolve-ADTErrorRecord -Property InnerException

        Enumerates only the InnerException property of the last ErrorRecord.

    .EXAMPLE
        Resolve-ADTErrorRecord -ExcludeErrorInvocation

        Enumerates the details of the last ErrorRecord, excluding the invocation information.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory = $false)]
        [SupportsWildcards()]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Property = ('Message', 'InnerException', 'FullyQualifiedErrorId', 'ScriptStackTrace', 'PositionMessage'),

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeErrorRecord,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeErrorInvocation,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeErrorException,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$ExcludeErrorInnerException
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $propsIsWildCard = $($Property).Equals('*')

        # Allows selecting and filtering the properties on the error object if they exist.
        filter Get-ErrorPropertyNames
        {
            [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
                [ValidateNotNullOrEmpty()]
                [System.Object]$InputObject
            )

            # Store all properties.
            $properties = $InputObject | & $Script:CommandTable.'Get-Member' -MemberType *Property | & $Script:CommandTable.'Select-Object' -ExpandProperty Name

            # If we've asked for all properties, return early with the above.
            if ($propsIsWildCard)
            {
                return $properties | & { process { if (![System.String]::IsNullOrWhiteSpace(($InputObject.$_ | & $Script:CommandTable.'Out-String').Trim())) { return $_ } } }
            }

            # Return all valid properties in the order used by the caller.
            return $Property | & { process { if (($properties -contains $_) -and ![System.String]::IsNullOrWhiteSpace(($InputObject.$_ | & $Script:CommandTable.'Out-String').Trim())) { return $_ } } }
        }
    }

    process
    {
        # Build out error objects to process in the right order.
        $errorObjects = $(
            $canDoException = !$ExcludeErrorException -and $ErrorRecord.Exception
            if (!$propsIsWildCard -and $canDoException)
            {
                $ErrorRecord.Exception
            }
            if (!$ExcludeErrorRecord)
            {
                $ErrorRecord
            }
            if (!$ExcludeErrorInvocation -and $ErrorRecord.InvocationInfo)
            {
                $ErrorRecord.InvocationInfo
            }
            if ($propsIsWildCard -and $canDoException)
            {
                $ErrorRecord.Exception
            }
        )

        # Open property collector and build it out.
        $logErrorProperties = [ordered]@{}
        foreach ($errorObject in $errorObjects)
        {
            # Store initial property count.
            $propCount = $logErrorProperties.Count

            # Add in all properties for the object.
            foreach ($propName in ($errorObject | Get-ErrorPropertyNames))
            {
                $logErrorProperties.Add($propName, ($errorObject.$propName).ToString().Trim())
            }

            # Append a new line to the last value for formatting purposes.
            if (!$propCount.Equals($logErrorProperties.Count))
            {
                $logErrorProperties.($logErrorProperties.Keys | & $Script:CommandTable.'Select-Object' -Last 1) += "`n"
            }
        }

        # Build out error properties.
        $logErrorMessage = [System.String]::Join("`n", "Error Record:", "-------------", $null, (& $Script:CommandTable.'Out-String' -InputObject (& $Script:CommandTable.'Format-List' -InputObject ([pscustomobject]$logErrorProperties)) -Width ([System.Int32]::MaxValue)).Trim())

        # Capture Error Inner Exception(s).
        if (!$ExcludeErrorInnerException -and $ErrorRecord.Exception -and $ErrorRecord.Exception.InnerException)
        {
            # Set up initial variables.
            $innerExceptions = [System.Collections.Specialized.StringCollection]::new()
            $errInnerException = $ErrorRecord.Exception.InnerException

            # Get all inner exceptions.
            while ($errInnerException)
            {
                # Add a divider if we've already added a record.
                if ($innerExceptions.Count)
                {
                    $null = $innerExceptions.Add("`n$('~' * 40)`n")
                }

                # Add error record and get next inner exception.
                $null = $innerExceptions.Add(($errInnerException | & $Script:CommandTable.'Select-Object' -Property ($errInnerException | Get-ErrorPropertyNames) | & $Script:CommandTable.'Format-List' | & $Script:CommandTable.'Out-String' -Width ([System.Int32]::MaxValue)).Trim())
                $errInnerException = $errInnerException.InnerException
            }

            # Output all inner exceptions to the caller.
            $logErrorMessage += "`n`n`n$([System.String]::Join("`n", "Error Inner Exception(s):", "-------------------------", $null, ($innerExceptions -join "`n")))"
        }

        # Output the error message to the caller.
        return $logErrorMessage
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Send-ADTKeys
#
#-----------------------------------------------------------------------------

function Send-ADTKeys
{
    <#
    .SYNOPSIS
        Send a sequence of keys to one or more application windows.

    .DESCRIPTION
        Send a sequence of keys to one or more application windows. If the window title searched for returns more than one window, then all of them will receive the sent keys.

        Function does not work in SYSTEM context unless launched with "psexec.exe -s -i" to run it as an interactive process under the SYSTEM account.

    .PARAMETER WindowTitle
        The title of the application window to search for using regex matching.

    .PARAMETER GetAllWindowTitles
        Get titles for all open windows on the system.

    .PARAMETER WindowHandle
        Send keys to a specific window where the Window Handle is already known.

    .PARAMETER Keys
        The sequence of keys to send. Info on Key input at: http://msdn.microsoft.com/en-us/library/System.Windows.Forms.SendKeys(v=vs.100).aspx

    .PARAMETER WaitSeconds
        An optional number of seconds to wait after the sending of the keys.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Send-ADTKeys -WindowTitle 'foobar - Notepad' -Keys 'Hello world'

        Send the sequence of keys "Hello world" to the application titled "foobar - Notepad".

    .EXAMPLE
        Send-ADTKeys -WindowTitle 'foobar - Notepad' -Keys 'Hello world' -WaitSeconds 5

        Send the sequence of keys "Hello world" to the application titled "foobar - Notepad" and wait 5 seconds.

    .EXAMPLE
        Send-ADTKeys -WindowHandle ([IntPtr]17368294) -Keys 'Hello World'

        Send the sequence of keys "Hello World" to the application with a Window Handle of '17368294'.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        http://msdn.microsoft.com/en-us/library/System.Windows.Forms.SendKeys(v=vs.100).aspx

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'WindowTitle')]
        [AllowEmptyString()]
        [ValidateNotNull()]
        [System.String]$WindowTitle,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'AllWindowTitles')]
        [System.Management.Automation.SwitchParameter]$GetAllWindowTitles,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'WindowHandle')]
        [ValidateNotNullOrEmpty()]
        [System.IntPtr]$WindowHandle,

        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'WindowTitle')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'AllWindowTitles')]
        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = 'WindowHandle')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Keys,

        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'WindowTitle')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'AllWindowTitles')]
        [Parameter(Mandatory = $false, Position = 4, ParameterSetName = 'WindowHandle')]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$WaitSeconds
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Internal worker filter.
        filter Send-ADTKeysToWindow
        {
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
                [ValidateNotNullOrEmpty()]
                [System.IntPtr]$WindowHandle,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.String]$Keys,

                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [System.Int32]$WaitSeconds
            )

            try
            {
                try
                {
                    # Bring the window to the foreground and make sure it's enabled.
                    if (![PSADT.GUI.UiAutomation]::BringWindowToFront($WindowHandle))
                    {
                        $naerParams = @{
                            Exception = [System.ApplicationException]::new('Failed to bring window to foreground.')
                            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                            ErrorId = 'WindowHandleForegroundError'
                            TargetObject = $WindowHandle
                            RecommendedAction = "Please check the status of this window and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                    if (![PSADT.LibraryInterfaces.User32]::IsWindowEnabled($WindowHandle))
                    {
                        $naerParams = @{
                            Exception = [System.ApplicationException]::new('Unable to send keys to window because it may be disabled due to a modal dialog being shown.')
                            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                            ErrorId = 'WindowHandleDisabledError'
                            TargetObject = $WindowHandle
                            RecommendedAction = "Please check the status of this window and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }

                    # Send the Key sequence.
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Sending key(s) [$Keys] to window title [$($Window.WindowTitle)] with window handle [$WindowHandle]."
                    [System.Windows.Forms.SendKeys]::SendWait($Keys)
                    if ($WaitSeconds)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Sleeping for [$WaitSeconds] seconds."
                        & $Script:CommandTable.'Start-Sleep' -Seconds $WaitSeconds
                    }
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to send keys to window title [$($Window.WindowTitle)] with window handle [$WindowHandle].`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
            }
        }

        # Set up parameter splat for worker filter.
        $sktwParams = @{ Keys = $Keys }; if ($PSBoundParameters.ContainsKey('Keys')) { $sktwParams.Add('WaitSeconds', $WaitSeconds) }
    }

    process
    {
        try
        {
            try
            {
                # Process the specified input.
                if ($WindowHandle)
                {
                    if (!($Window = & $Script:CommandTable.'Get-ADTWindowTitle' -GetAllWindowTitles | & { process { if ($_.WindowHandle -eq $WindowHandle) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "No windows with Window Handle [$WindowHandle] were discovered." -Severity 2
                        return
                    }
                    Send-ADTKeysToWindow -WindowHandle $Window.WindowHandle @sktwParams
                }
                else
                {
                    if (!($AllWindows = if ($GetAllWindowTitles) { & $Script:CommandTable.'Get-ADTWindowTitle' -GetAllWindowTitles $GetAllWindowTitles } else { & $Script:CommandTable.'Get-ADTWindowTitle' -WindowTitle $WindowTitle } ))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'No windows with the specified details were discovered.' -Severity 2
                        return
                    }
                    $AllWindows | Send-ADTKeysToWindow @sktwParams
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to send keys to specified window."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTActiveSetup
#
#-----------------------------------------------------------------------------

function Set-ADTActiveSetup
{
    <#
    .SYNOPSIS
        Creates an Active Setup entry in the registry to execute a file for each user upon login.

    .DESCRIPTION
        Active Setup allows handling of per-user changes registry/file changes upon login.

        A registry key is created in the HKLM registry hive which gets replicated to the HKCU hive when a user logs in.

        If the "Version" value of the Active Setup entry in HKLM is higher than the version value in HKCU, the file referenced in "StubPath" is executed.

        This Function:
            - Creates the registry entries in "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$($adtSession.InstallName)".
            - Creates StubPath value depending on the file extension of the $StubExePath parameter.
            - Handles Version value with YYYYMMDDHHMMSS granularity to permit re-installs on the same day and still trigger Active Setup after Version increase.
            - Copies/overwrites the StubPath file to $StubExePath destination path if file exists in 'Files' subdirectory of script directory.
            - Executes the StubPath file for the current user based on $NoExecuteForCurrentUser (no need to logout/login to trigger Active Setup).

    .PARAMETER StubExePath
        Use this parameter to specify the destination path of the file that will be executed upon user login.

        Note: Place the file you want users to execute in the '\Files' subdirectory of the script directory and the toolkit will install it to the path specificed in this parameter.

    .PARAMETER Arguments
        Arguments to pass to the file being executed.

    .PARAMETER Wow6432Node
        Specify this switch to use Active Setup entry under Wow6432Node on a 64-bit OS. Default is: $false.

    .PARAMETER ExecutionPolicy
        Specifies the ExecutionPolicy to set when StubExePath is a PowerShell script. Default is: system's ExecutionPolicy.

    .PARAMETER Version
        Optional. Specify version for Active setup entry. Active Setup is not triggered if Version value has more than 8 consecutive digits. Use commas to get around this limitation. Default: YYYYMMDDHHMMSS

        Note:
            - Do not use this parameter if it is not necessary. PSADT will handle this parameter automatically using the time of the installation as the version number.
            - In Windows 10, Scripts and EXEs might be blocked by AppLocker. Ensure that the path given to -StubExePath will permit end users to run Scripts and EXEs unelevated.

    .PARAMETER Locale
        Optional. Arbitrary string used to specify the installation language of the file being executed. Not replicated to HKCU.

    .PARAMETER PurgeActiveSetupKey
        Remove Active Setup entry from HKLM registry hive. Will also load each logon user's HKCU registry hive to remove Active Setup entry. Function returns after purging.

    .PARAMETER DisableActiveSetup
        Disables the Active Setup entry so that the StubPath file will not be executed. This also enables -NoExecuteForCurrentUser.

    .PARAMETER NoExecuteForCurrentUser
        Specifies whether the StubExePath should be executed for the current user. Since this user is already logged in, the user won't have the application started without logging out and logging back in.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if Active Setup entry was created or updated, $false if Active Setup entry was not created or updated.

    .EXAMPLE
        Set-ADTActiveSetup -StubExePath 'C:\Users\Public\Company\ProgramUserConfig.vbs' -Arguments '/Silent' -Description 'Program User Config' -Key 'ProgramUserConfig' -Locale 'en'

    .EXAMPLE
        Set-ADTActiveSetup -StubExePath "$envWinDir\regedit.exe" -Arguments "/S `"%SystemDrive%\Program Files (x86)\PS App Deploy\PSAppDeployHKCUSettings.reg`"" -Description 'PS App Deploy Config' -Key 'PS_App_Deploy_Config'

    .EXAMPLE
        Set-ADTActiveSetup -Key 'ProgramUserConfig' -PurgeActiveSetupKey

        Delete "ProgramUserConfig" active setup entry from all registry hives.

    .NOTES
        An active ADT session is NOT required to use this function.

        Original code borrowed from: Denis St-Pierre (Ottawa, Canada), Todd MacNaught (Ottawa, Canada)

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(DefaultParameterSetName = 'Create')]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Create')]
        [ValidateScript({
                if (('.exe', '.vbs', '.cmd', '.bat', '.ps1', '.js') -notcontains ($StubExeExt = [System.IO.Path]::GetExtension($_)))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName StubExePath -ProvidedValue $_ -ExceptionMessage "Unsupported Active Setup StubPath file extension [$StubExeExt]."))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$StubExePath,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Arguments,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [Microsoft.PowerShell.ExecutionPolicy]$ExecutionPolicy,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Version = ((& $Script:CommandTable.'Get-Date' -Format 'yyMM,ddHH,mmss').ToString()), # Ex: 1405,1515,0522

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Locale,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [System.Management.Automation.SwitchParameter]$DisableActiveSetup,

        [Parameter(Mandatory = $true, ParameterSetName = 'Purge')]
        [System.Management.Automation.SwitchParameter]$PurgeActiveSetupKey,

        [Parameter(Mandatory = $false, ParameterSetName = 'Create')]
        [System.Management.Automation.SwitchParameter]$NoExecuteForCurrentUser
    )

    dynamicparam
    {
        # Attempt to get the most recent ADTSession object.
        $adtSession = if (& $Script:CommandTable.'Test-ADTSessionActive')
        {
            & $Script:CommandTable.'Get-ADTSession'
        }

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('Key', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Key', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Name of the registry key for the Active Setup entry. Defaults to active session InstallName.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('Description', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Description', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Description for the Active Setup. Users will see "Setting up personalized settings for: $Description" at logon. Defaults to active session InstallName.'; ParameterSetName = 'Create' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Set defaults for when there's an active ADTSession and overriding values haven't been specified.
        $Description = if ($PSCmdlet.ParameterSetName.Equals('Create'))
        {
            if (!$PSBoundParameters.ContainsKey('Description'))
            {
                $adtSession.InstallName
            }
            else
            {
                $PSBoundParameters.Description
            }
        }
        $Key = if (!$PSBoundParameters.ContainsKey('Key'))
        {
            $adtSession.InstallName
        }
        else
        {
            $PSBoundParameters.Key
        }

        # Define initial variables.
        $runAsActiveUser = & $Script:CommandTable.'Get-ADTRunAsActiveUser'
        $CUStubExePath = $null
        $CUArguments = $null
        $StubExeExt = [System.IO.Path]::GetExtension($StubExePath)
        $StubPath = $null

        # Define internal function to test current ActiveSetup stuff.
        function Test-ADTActiveSetup
        {
            [CmdletBinding()]
            [OutputType([System.Boolean])]
            param
            (
                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.String]$HKLMKey,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.String]$HKCUKey,

                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [System.String]$SID
            )

            # Set up initial variables.
            $HKCUProps = if ($SID)
            {
                & $Script:CommandTable.'Get-ADTRegistryKey' -Key $HKCUKey -SID $SID
            }
            else
            {
                & $Script:CommandTable.'Get-ADTRegistryKey' -Key $HKCUKey
            }
            $HKLMProps = & $Script:CommandTable.'Get-ADTRegistryKey' -Key $HKLMKey
            $HKCUVer = $HKCUProps | & $Script:CommandTable.'Select-Object' -ExpandProperty Version -ErrorAction Ignore
            $HKLMVer = $HKLMProps | & $Script:CommandTable.'Select-Object' -ExpandProperty Version -ErrorAction Ignore
            $HKLMInst = $HKLMProps | & $Script:CommandTable.'Select-Object' -ExpandProperty IsInstalled -ErrorAction Ignore

            # HKLM entry not present. Nothing to run.
            if (!$HKLMProps)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM active setup entry is not present.'
                return $false
            }

            # HKLM entry present, but disabled. Nothing to run.
            if ($HKLMInst -eq 0)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM active setup entry is present, but it is disabled (IsInstalled set to 0).'
                return $false
            }

            # HKLM entry present and HKCU entry is not. Run the StubPath.
            if (!$HKCUProps)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM active setup entry is present. HKCU active setup entry is not present.'
                return $true
            }

            # Both entries present. HKLM entry does not have Version property. Nothing to run.
            if (!$HKLMVer)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. HKLM Version property is missing.'
                return $false
            }

            # Both entries present. HKLM entry has Version property, but HKCU entry does not. Run the StubPath.
            if (!$HKCUVer)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. HKCU Version property is missing.'
                return $true
            }

            # After cleanup, the HKLM Version property is empty. Considering it missing. HKCU is present so nothing to run.
            if (!($HKLMValidVer = [System.String]::Join($null, ($HKLMVer.GetEnumerator() | & { process { if ([System.Char]::IsDigit($_) -or ($_ -eq ',')) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1))))
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. HKLM Version property is invalid.'
                return $false
            }

            # After cleanup, the HKCU Version property is empty while HKLM Version property is not. Run the StubPath.
            if (!($HKCUValidVer = [System.String]::Join($null, ($HKCUVer.GetEnumerator() | & { process { if ([System.Char]::IsDigit($_) -or ($_ -eq ',')) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1))))
            {
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. HKCU Version property is invalid.'
                return $true
            }

            # Both entries present, with a Version property. Compare the Versions.
            try
            {
                # Convert the version property to Version type and compare.
                if (([System.Version]$HKLMValidVer.Replace(',', '.')) -gt ([System.Version]$HKCUValidVer.Replace(',', '.')))
                {
                    # HKLM is greater, run the StubPath.
                    & $Script:CommandTable.'Write-ADTLogEntry' "HKLM and HKCU active setup entries are present. Both contain Version properties, and the HKLM Version is greater."
                    return $true
                }
                else
                {
                    # The HKCU version is equal or higher than HKLM version, Nothing to run.
                    & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. Both contain Version properties. However, they are either the same or the HKCU Version property is higher.'
                    return $false
                }
            }
            catch
            {
                # Failed to convert version property to Version type.
                $null = $null
            }

            # Check whether the Versions were split into the same number of strings. Split the version by commas.
            if (($SplitHKLMValidVer = $HKLMValidVer.Split(',')).Count -ne ($SplitHKCUValidVer = $HKCUValidVer.Split(',')).Count)
            {
                # The versions are different length - more commas
                if ($SplitHKLMValidVer.Count -gt $SplitHKCUValidVer.Count)
                {
                    # HKLM is longer, Run the StubPath.
                    & $Script:CommandTable.'Write-ADTLogEntry' "HKLM and HKCU active setup entries are present. Both contain Version properties. However, the HKLM Version has more version fields."
                    return $true
                }
                else
                {
                    # HKCU is longer, Nothing to run.
                    & $Script:CommandTable.'Write-ADTLogEntry' "HKLM and HKCU active setup entries are present. Both contain Version properties. However, the HKCU Version has more version fields."
                    return $false
                }
            }

            # The Versions have the same number of strings. Compare them
            try
            {
                for ($i = 0; $i -lt $SplitHKLMValidVer.Count; $i++)
                {
                    # Parse the version is UINT64.
                    if ([UInt64]::Parse($SplitHKCUValidVer[$i]) -lt [UInt64]::Parse($SplitHKLMValidVer[$i]))
                    {
                        # The HKCU ver is lower, Run the StubPath.
                        & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. Both Version properties are present and valid. However, HKCU Version property is lower.'
                        return $true
                    }
                }
                # The HKCU version is equal or higher than HKLM version, Nothing to run.
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. Both Version properties are present and valid. However, they are either the same or HKCU Version property is higher.'
                return $false
            }
            catch
            {
                # Failed to parse strings as UInt64, Run the StubPath.
                & $Script:CommandTable.'Write-ADTLogEntry' 'HKLM and HKCU active setup entries are present. Both Version properties are present and valid. However, parsing string numerics to 64-bit integers failed.' -Severity 2
                return $true
            }
        }

        # Define internal function to the required ActiveSetup registry keys.
        function Set-ADTActiveSetupRegistryEntry
        {
            [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'This is an internal worker function that requires no end user confirmation.')]
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.String]$RegPath,

                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [System.String]$SID,

                [Parameter(Mandatory = $false)]
                [ValidateNotNullOrEmpty()]
                [System.String]$Version,

                [Parameter(Mandatory = $false)]
                [AllowEmptyString()]
                [System.String]$Locale,

                [Parameter(Mandatory = $false)]
                [System.Management.Automation.SwitchParameter]$DisableActiveSetup
            )

            $srkParams = if ($SID) { @{ SID = $SID } } else { @{} }
            & $Script:CommandTable.'Set-ADTRegistryKey' -Key $RegPath -Name '(Default)' -Value $Description @srkParams
            & $Script:CommandTable.'Set-ADTRegistryKey' -Key $RegPath -Name 'Version' -Value $Version @srkParams
            & $Script:CommandTable.'Set-ADTRegistryKey' -Key $RegPath -Name 'StubPath' -Value $StubPath -Type 'String' @srkParams
            if (![System.String]::IsNullOrWhiteSpace($Locale))
            {
                & $Script:CommandTable.'Set-ADTRegistryKey' -Key $RegPath -Name 'Locale' -Value $Locale @srkParams
            }

            # Only Add IsInstalled to HKLM.
            if ($RegPath.Contains('HKEY_LOCAL_MACHINE'))
            {
                & $Script:CommandTable.'Set-ADTRegistryKey' -Key $RegPath -Name 'IsInstalled' -Value ([System.UInt32]!$DisableActiveSetup) -Type 'DWord' @srkParams
            }
        }
    }

    process
    {
        try
        {
            try
            {
                # Set up the relevant keys, factoring in bitness and architecture.
                if ($Wow6432Node -and [System.Environment]::Is64BitOperatingSystem)
                {
                    $HKLMRegKey = "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$Key"
                    $HKCURegKey = "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\$Key"
                }
                else
                {
                    $HKLMRegKey = "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\$Key"
                    $HKCURegKey = "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Active Setup\Installed Components\$Key"
                }

                # Delete Active Setup registry entry from the HKLM hive and for all logon user registry hives on the system.
                if ($PurgeActiveSetupKey)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing Active Setup entry [$HKLMRegKey]."
                    & $Script:CommandTable.'Remove-ADTRegistryKey' -Key $HKLMRegKey -Recurse

                    if ($runAsActiveUser)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing Active Setup entry [$HKCURegKey] for all logged on user registry hives on the system."
                        & $Script:CommandTable.'Invoke-ADTAllUsersRegistryAction' -UserProfiles (& $Script:CommandTable.'Get-ADTUserProfiles' -ExcludeDefaultUser | & { process { if ($_.SID -eq $runAsActiveUser.SID) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1) -ScriptBlock {
                            if (& $Script:CommandTable.'Get-ADTRegistryKey' -Key $HKCURegKey -SID $_.SID)
                            {
                                & $Script:CommandTable.'Remove-ADTRegistryKey' -Key $HKCURegKey -SID $_.SID -Recurse
                            }
                        }
                    }
                    return
                }

                # Copy file to $StubExePath from the 'Files' subdirectory of the script directory (if it exists there).
                $StubExePath = [System.Environment]::ExpandEnvironmentVariables($StubExePath)
                if ($adtSession -and $adtSession.DirFiles)
                {
                    $StubExeFile = & $Script:CommandTable.'Join-Path' -Path $adtSession.DirFiles -ChildPath ($ActiveSetupFileName = [System.IO.Path]::GetFileName($StubExePath))
                    if (& $Script:CommandTable.'Test-Path' -LiteralPath $StubExeFile -PathType Leaf)
                    {
                        # This will overwrite the StubPath file if $StubExePath already exists on target.
                        & $Script:CommandTable.'Copy-ADTFile' -Path $StubExeFile -Destination $StubExePath -ErrorAction Stop
                    }
                }

                # Check if the $StubExePath file exists.
                if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $StubExePath -PathType Leaf))
                {
                    $naerParams = @{
                        Exception = [System.IO.FileNotFoundException]::new("Active Setup StubPath file [$ActiveSetupFileName] is missing.")
                        Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                        ErrorId = 'ActiveSetupFileNotFound'
                        TargetObject = $ActiveSetupFileName
                        RecommendedAction = "Please confirm the provided value and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Define Active Setup StubPath according to file extension of $StubExePath.
                switch ($StubExeExt)
                {
                    '.exe'
                    {
                        $CUStubExePath = $StubExePath
                        $CUArguments = $Arguments
                        $StubPath = if ([System.String]::IsNullOrWhiteSpace($Arguments))
                        {
                            "`"$CUStubExePath`""
                        }
                        else
                        {
                            "`"$CUStubExePath`" $CUArguments"
                        }
                        break
                    }
                    { $_ -in '.js', '.vbs' }
                    {
                        $CUStubExePath = "$([System.Environment]::SystemDirectory)\wscript.exe"
                        $CUArguments = if ([System.String]::IsNullOrWhiteSpace($Arguments))
                        {
                            "//nologo `"$StubExePath`""
                        }
                        else
                        {
                            "//nologo `"$StubExePath`"  $Arguments"
                        }
                        $StubPath = "`"$CUStubExePath`" $CUArguments"
                        break
                    }
                    { $_ -in '.cmd', '.bat' }
                    {
                        $CUStubExePath = "$([System.Environment]::SystemDirectory)\cmd.exe"
                        # Prefix any CMD.exe metacharacters ^ or & with ^ to escape them - parentheses only require escaping when there's no space in the path!
                        $StubExePath = if ($StubExePath.Trim() -match '\s')
                        {
                            $StubExePath -replace '([&^])', '^$1'
                        }
                        else
                        {
                            $StubExePath -replace '([()&^])', '^$1'
                        }
                        $CUArguments = if ([System.String]::IsNullOrWhiteSpace($Arguments))
                        {
                            "/C `"$StubExePath`""
                        }
                        else
                        {
                            "/C `"`"$StubExePath`" $Arguments`""
                        }
                        $StubPath = "`"$CUStubExePath`" $CUArguments"
                        break
                    }
                    '.ps1'
                    {
                        $CUStubExePath = & $Script:CommandTable.'Get-ADTPowerShellProcessPath'
                        $CUArguments = if ([System.String]::IsNullOrWhiteSpace($Arguments))
                        {
                            "$(if ($PSBoundParameters.ContainsKey('ExecutionPolicy')) { "-ExecutionPolicy $ExecutionPolicy" })-NoProfile -NoLogo -WindowStyle Hidden -File `"$StubExePath`""
                        }
                        else
                        {
                            "$(if ($PSBoundParameters.ContainsKey('ExecutionPolicy')) { "-ExecutionPolicy $ExecutionPolicy" })-NoProfile -NoLogo -WindowStyle Hidden -File `"$StubExePath`" $Arguments"
                        }
                        $StubPath = "`"$CUStubExePath`" $CUArguments"
                        break
                    }
                }

                # Define common parameters split for Set-ADTActiveSetupRegistryEntry.
                $sasreParams = @{
                    Version = $Version
                    Locale = $Locale
                    DisableActiveSetup = $DisableActiveSetup
                }

                # Create the Active Setup entry in the registry.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Adding Active Setup Key for local machine: [$HKLMRegKey]."
                Set-ADTActiveSetupRegistryEntry @sasreParams -RegPath $HKLMRegKey

                # Execute the StubPath file for the current user as long as not in Session 0.
                if ($NoExecuteForCurrentUser)
                {
                    return
                }

                if (![System.Diagnostics.Process]::GetCurrentProcess().SessionId)
                {
                    if (!$runAsActiveUser)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Session 0 detected: No logged in users detected. Active Setup StubPath file will execute when users first log into their account.'
                        return
                    }

                    # Skip if Active Setup reg key is present and Version is equal or higher
                    if (!(Test-ADTActiveSetup -HKLMKey $HKLMRegKey -HKCUKey $HKCURegKey -SID $runAsActiveUser.SID))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Session 0 detected: Skipping executing Active Setup StubPath file for currently logged in user [$($runAsActiveUser.NTAccount)]." -Severity 2
                        return
                    }

                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Session 0 detected: Executing Active Setup StubPath file for currently logged in user [$($runAsActiveUser.NTAccount)]."
                    if ($CUArguments)
                    {
                        & $Script:CommandTable.'Start-ADTProcessAsUser' -FilePath $CUStubExePath -ArgumentList $CUArguments -Wait -HideWindow
                    }
                    else
                    {
                        & $Script:CommandTable.'Start-ADTProcessAsUser' -FilePath $CUStubExePath -Wait -HideWindow
                    }

                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Adding Active Setup Key for the current user: [$HKCURegKey]."
                    Set-ADTActiveSetupRegistryEntry @sasreParams -RegPath $HKCURegKey -SID $runAsActiveUser.SID
                }
                else
                {
                    # Skip if Active Setup reg key is present and Version is equal or higher
                    if (!(Test-ADTActiveSetup -HKLMKey $HKLMRegKey -HKCUKey $HKCURegKey))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Skipping executing Active Setup StubPath file for current user.' -Severity 2
                        return
                    }

                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Executing Active Setup StubPath file for the current user.'
                    if ($CUArguments)
                    {
                        & $Script:CommandTable.'Start-ADTProcess' -FilePath $CUStubExePath -ArgumentList $CUArguments
                    }
                    else
                    {
                        & $Script:CommandTable.'Start-ADTProcess' -FilePath $CUStubExePath
                    }

                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Adding Active Setup Key for the current user: [$HKCURegKey]."
                    Set-ADTActiveSetupRegistryEntry @sasreParams -RegPath $HKCURegKey
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to set Active Setup registry entry."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTDeferHistory
#
#-----------------------------------------------------------------------------

function Set-ADTDeferHistory
{
    <#
    .SYNOPSIS
        Set the history of deferrals in the registry for the current application.

    .DESCRIPTION
        Set the history of deferrals in the registry for the current application.

    .PARAMETER DeferTimesRemaining
        Specify the number of deferrals remaining.

    .PARAMETER DeferDeadline
        Specify the deadline for the deferral.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Set-DeferHistory

    .NOTES
        An active ADT session is required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com

    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$DeferTimesRemaining,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [System.String]$DeferDeadline
    )

    try
    {
        (& $Script:CommandTable.'Get-ADTSession').SetDeferHistory($(if ($PSBoundParameters.ContainsKey('DeferTimesRemaining')) { $DeferTimesRemaining }), $DeferDeadline)
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTIniValue
#
#-----------------------------------------------------------------------------

function Set-ADTIniValue
{
    <#
    .SYNOPSIS
        Opens an INI file and sets the value of the specified section and key.

    .DESCRIPTION
        Opens an INI file and sets the value of the specified section and key. If the value is set to $null, the key will be removed from the section.

    .PARAMETER FilePath
        Path to the INI file.

    .PARAMETER Section
        Section within the INI file.

    .PARAMETER Key
        Key within the section of the INI file.

    .PARAMETER Value
        Value for the key within the section of the INI file. To remove a value, set this variable to $null.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Set-ADTIniValue -FilePath "$env:ProgramFilesX86\IBM\Notes\notes.ini" -Section 'Notes' -Key 'KeyFileName' -Value 'MyFile.ID'

        Sets the 'KeyFileName' key in the 'Notes' section of the 'notes.ini' file to 'MyFile.ID'.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
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
        [System.Object]$Value
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Writing INI Key Value: [Section = $Section] [Key = $Key] [Value = $Value]."
        try
        {
            try
            {
                [PSADT.Configuration.IniFile]::WriteSectionKeyValue($Section, $Key, $Value, $FilePath)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to write INI file key value."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTItemPermission
#
#-----------------------------------------------------------------------------

function Set-ADTItemPermission
{
    <#
    .SYNOPSIS
        Allows you to easily change permissions on files or folders.

    .DESCRIPTION
        Allows you to easily change permissions on files or folders for a given user or group. You can add, remove or replace permissions, set inheritance and propagation.

    .PARAMETER Path
        Path to the folder or file you want to modify (ex: C:\Temp)

    .PARAMETER User
        One or more user names (ex: BUILTIN\Users, DOMAIN\Admin) to give the permissions to. If you want to use SID, prefix it with an asterisk * (ex: *S-1-5-18)

    .PARAMETER Permission
        Permission or list of permissions to be set/added/removed/replaced. To see all the possible permissions go to 'http://technet.microsoft.com/fr-fr/library/ff730951.aspx'.

        Permission DeleteSubdirectoriesAndFiles does not apply to files.

    .PARAMETER PermissionType
        Sets Access Control Type of the permissions. Allowed options: Allow, Deny

    .PARAMETER Inheritance
        Sets permission inheritance. Does not apply to files. Multiple options can be specified. Allowed options: ObjectInherit, ContainerInherit, None

        None - The permission entry is not inherited by child objects, ObjectInherit - The permission entry is inherited by child leaf objects. ContainerInherit - The permission entry is inherited by child container objects.

    .PARAMETER Propagation
        Sets how to propagate inheritance. Does not apply to files. Allowed options: None, InheritOnly, NoPropagateInherit

        None - Specifies that no inheritance flags are set. NoPropagateInherit - Specifies that the permission entry is not propagated to child objects. InheritOnly - Specifies that the permission entry is propagated only to child objects. This includes both container and leaf child objects.

    .PARAMETER Method
        Specifies which method will be used to apply the permissions. Allowed options: Add, Set, Reset.

        Add - adds permissions rules but it does not remove previous permissions, Set - overwrites matching permission rules with new ones, Reset - removes matching permissions rules and then adds permission rules, Remove - Removes matching permission rules, RemoveSpecific - Removes specific permissions, RemoveAll - Removes all permission rules for specified user/s

    .PARAMETER EnableInheritance
        Enables inheritance on the files/folders.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Set-ADTItemPermission -Path 'C:\Temp' -User 'DOMAIN\John', 'BUILTIN\Users' -Permission FullControl -Inheritance ObjectInherit,ContainerInherit

        Will grant FullControl permissions to 'John' and 'Users' on 'C:\Temp' and its files and folders children.

    .EXAMPLE
        Set-ADTItemPermission -Path 'C:\Temp\pic.png' -User 'DOMAIN\John' -Permission 'Read'

        Will grant Read permissions to 'John' on 'C:\Temp\pic.png'.

    .EXAMPLE
        Set-ADTItemPermission -Path 'C:\Temp\Private' -User 'DOMAIN\John' -Permission 'None' -Method 'RemoveAll'

        Will remove all permissions to 'John' on 'C:\Temp\Private'.

    .NOTES
        An active ADT session is NOT required to use this function.

        Original Author: Julian DA CUNHA - dacunha.julian@gmail.com, used with permission.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'PermissionType', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Method', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Path to the folder or file you want to modify (ex: C:\Temp)', ParameterSetName = 'DisableInheritance')]
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Path to the folder or file you want to modify (ex: C:\Temp)', ParameterSetName = 'EnableInheritance')]
        [ValidateScript({
                if (!(& $Script:CommandTable.'Test-Path' -Path $_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [Alias('File', 'Folder')]
        [System.String]$Path,

        [Parameter(Mandatory = $true, Position = 1, HelpMessage = 'One or more user names (ex: BUILTIN\Users, DOMAIN\Admin). If you want to use SID, prefix it with an asterisk * (ex: *S-1-5-18)', ParameterSetName = 'DisableInheritance')]
        [Alias('Username', 'Users', 'SID', 'Usernames')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$User,

        [Parameter(Mandatory = $true, Position = 2, HelpMessage = "Permission or list of permissions to be set/added/removed/replaced. To see all the possible permissions go to 'http://technet.microsoft.com/fr-fr/library/ff730951.aspx'", ParameterSetName = 'DisableInheritance')]
        [Alias('Acl', 'Grant', 'Permissions', 'Deny')]
        [ValidateNotNullOrEmpty()]
        [System.Security.AccessControl.FileSystemRights]$Permission,

        [Parameter(Mandatory = $false, Position = 3, HelpMessage = 'Whether you want to set Allow or Deny permissions', ParameterSetName = 'DisableInheritance')]
        [Alias('AccessControlType')]
        [ValidateNotNullOrEmpty()]
        [System.Security.AccessControl.AccessControlType]$PermissionType = [System.Security.AccessControl.AccessControlType]::Allow,

        [Parameter(Mandatory = $false, Position = 4, HelpMessage = 'Sets how permissions are inherited', ParameterSetName = 'DisableInheritance')]
        [ValidateNotNullOrEmpty()]
        [System.Security.AccessControl.InheritanceFlags]$Inheritance = [System.Security.AccessControl.InheritanceFlags]::None,

        [Parameter(Mandatory = $false, Position = 5, HelpMessage = 'Sets how to propage inheritance flags', ParameterSetName = 'DisableInheritance')]
        [ValidateNotNullOrEmpty()]
        [System.Security.AccessControl.PropagationFlags]$Propagation = [System.Security.AccessControl.PropagationFlags]::None,

        [Parameter(Mandatory = $false, Position = 6, HelpMessage = 'Specifies which method will be used to add/remove/replace permissions.', ParameterSetName = 'DisableInheritance')]
        [ValidateSet('AddAccessRule', 'SetAccessRule', 'ResetAccessRule', 'RemoveAccessRule', 'RemoveAccessRuleSpecific', 'RemoveAccessRuleAll')]
        [Alias('ApplyMethod', 'ApplicationMethod')]
        [System.String]$Method = 'AddAccessRule',

        [Parameter(Mandatory = $true, Position = 1, HelpMessage = 'Enables inheritance, which removes explicit permissions.', ParameterSetName = 'EnableInheritance')]
        [System.Management.Automation.SwitchParameter]$EnableInheritance
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Get object ACLs and enable inheritance.
                if ($EnableInheritance)
                {
                    ($Acl = & $Script:CommandTable.'Get-Acl' -Path $Path).SetAccessRuleProtection($false, $true)
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Enabling Inheritance on path [$Path]."
                    $null = & $Script:CommandTable.'Set-Acl' -Path $Path -AclObject $Acl
                    return
                }

                # Modify variables to remove file incompatible flags if this is a file.
                if (& $Script:CommandTable.'Test-Path' -LiteralPath $Path -PathType Leaf)
                {
                    $Permission = $Permission -band (-bnot [System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles)
                    $Inheritance = [System.Security.AccessControl.InheritanceFlags]::None
                    $Propagation = [System.Security.AccessControl.PropagationFlags]::None
                }

                # Get object ACLs, disable inheritance but preserve inherited permissions.
                ($Acl = & $Script:CommandTable.'Get-Acl' -Path $Path).SetAccessRuleProtection($true, $true)
                $null = & $Script:CommandTable.'Set-Acl' -Path $Path -AclObject $Acl

                # Get updated ACLs - without inheritance.
                $Acl = & $Script:CommandTable.'Get-Acl' -Path $Path

                # Apply permissions on each user.
                $User.Trim() | & {
                    process
                    {
                        # Return early if the string is empty.
                        if (!$_.Length)
                        {
                            return
                        }

                        # Set Username.
                        [System.Security.Principal.NTAccount]$Username = if ($_.StartsWith('*'))
                        {
                            try
                            {
                                # Translate the SID.
                                & $Script:CommandTable.'ConvertTo-ADTNTAccountOrSID' -SID ($sid = $_.Remove(0, 1))
                            }
                            catch
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' "Failed to translate SID [$sid]. Skipping..." -Severity 2
                                continue
                            }
                        }
                        else
                        {
                            $_
                        }

                        # Set/Add/Remove/Replace permissions and log the changes.
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Changing permissions [Permissions:$Permission, InheritanceFlags:$Inheritance, PropagationFlags:$Propagation, AccessControlType:$PermissionType, Method:$Method] on path [$Path] for user [$Username]."
                        $Acl.$Method([System.Security.AccessControl.FileSystemAccessRule]::new($Username, $Permission, $Inheritance, $Propagation, $PermissionType))
                    }
                }

                # Use the prepared ACL.
                $null = & $Script:CommandTable.'Set-Acl' -Path $Path -AclObject $Acl
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTMsiProperty
#
#-----------------------------------------------------------------------------

function Set-ADTMsiProperty
{
    <#
    .SYNOPSIS
        Set a property in the MSI property table.

    .DESCRIPTION
        Set a property in the MSI property table.

    .PARAMETER Database
        Specify a ComObject representing an MSI database opened in view/modify/update mode.

    .PARAMETER PropertyName
        The name of the property to be set/modified.

    .PARAMETER PropertyValue
        The value of the property to be set/modified.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Set-ADTMsiProperty -Database $TempMsiPathDatabase -PropertyName 'ALLUSERS' -PropertyValue '1'

    .NOTES
        An active ADT session is NOT required to use this function.

        Original Author: Julian DA CUNHA - dacunha.julian@gmail.com, used with permission.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.__ComObject]$Database,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$PropertyName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$PropertyValue
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Setting the MSI Property Name [$PropertyName] with Property Value [$PropertyValue]."
        try
        {
            try
            {
                # Open the requested table view from the database.
                $View = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Database -MethodName OpenView -ArgumentList @("SELECT * FROM Property WHERE Property='$PropertyName'")
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Execute

                # Retrieve the requested property from the requested table and close off the view.
                # https://msdn.microsoft.com/en-us/library/windows/desktop/aa371136(v=vs.85).aspx
                $Record = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Fetch
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Close -ArgumentList @()
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($View)

                # Set the MSI property.
                $View = if ($Record)
                {
                    # If the property already exists, then create the view for updating the property.
                    & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Database -MethodName OpenView -ArgumentList @("UPDATE Property SET Value='$PropertyValue' WHERE Property='$PropertyName'")
                }
                else
                {
                    # If property does not exist, then create view for inserting the property.
                    & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Database -MethodName OpenView -ArgumentList @("INSERT INTO Property (Property, Value) VALUES ('$PropertyName','$PropertyValue')")
                }
                $null = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Execute
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to set the MSI Property Name [$PropertyName] with Property Value [$PropertyValue]."
        }
        finally
        {
            $null = try
            {
                if (& $Script:CommandTable.'Test-Path' -LiteralPath Microsoft.PowerShell.Core\Variable::View)
                {
                    & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $View -MethodName Close -ArgumentList @()
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($View)
                }
            }
            catch
            {
                $null
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTPowerShellCulture
#
#-----------------------------------------------------------------------------

function Set-ADTPowerShellCulture
{
    <#
    .SYNOPSIS
        Changes the current thread's Culture and UICulture to the specified culture.

    .DESCRIPTION
        This function changes the current thread's Culture and UICulture to the specified culture.

    .PARAMETER CultureInfo
        The culture to set the current thread's Culture and UICulture to. Can be a CultureInfo object, or any valid IETF BCP 47 language tag.

    .EXAMPLE
        Set-ADTPowerShellCulture -Culture en-US

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Globalization.CultureInfo]$CultureInfo
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $smaCultureResolver = [System.Reflection.Assembly]::Load('System.Management.Automation').GetType('Microsoft.PowerShell.NativeCultureResolver')
        $smaResolverFlags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
        [System.Globalization.CultureInfo[]]$validCultures = (& $Script:CommandTable.'Get-WinUserLanguageList').LanguageTag
    }

    process
    {
        try
        {
            try
            {
                # Test that the specified culture is installed or not.
                if (!$validCultures.Contains($CultureInfo))
                {
                    $naerParams = @{
                        Exception = [System.ArgumentException]::new("The language pack for [$CultureInfo] is not installed on this system.", $CultureInfo)
                        Category = [System.Management.Automation.ErrorCategory]::InvalidArgument
                        ErrorId = 'CultureNotInstalled'
                        TargetObject = $validCultures
                        RecommendedAction = "Please review the installed cultures within this error's TargetObject and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Reflectively update the culture to the specified value.
                # This will change PowerShell, but not its default variables like $PSCulture and $PSUICulture.
                $smaCultureResolver.GetField('m_Culture', $smaResolverFlags).SetValue($null, $CultureInfo)
                $smaCultureResolver.GetField('m_uiCulture', $smaResolverFlags).SetValue($null, $CultureInfo)
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTRegistryKey
#
#-----------------------------------------------------------------------------

function Set-ADTRegistryKey
{
    <#
    .SYNOPSIS
        Creates or sets a registry key name, value, and value data.

    .DESCRIPTION
        Creates a registry key name, value, and value data; it sets the same if it already exists. This function can also handle registry keys for specific user SIDs and 32-bit registry on 64-bit systems.

    .PARAMETER Key
        The registry key path.

    .PARAMETER Name
        The value name.

    .PARAMETER Value
        The value data.

    .PARAMETER Type
        The type of registry value to create or set. Options: 'Binary','DWord','ExpandString','MultiString','None','QWord','String','Unknown'. Default: String.

        DWord should be specified as a decimal.

    .PARAMETER Wow6432Node
        Specify this switch to write to the 32-bit registry (Wow6432Node) on 64-bit systems.

    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.

        Specify this parameter from the Invoke-ADTAllUsersRegistryAction function to read/edit HKCU registry settings for all users on the system.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Set-ADTRegistryKey -Key $blockedAppPath -Name 'Debugger' -Value $blockedAppDebuggerValue

        Creates or sets the 'Debugger' value in the specified registry key.

    .EXAMPLE
        Set-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE' -Name 'Application' -Type 'DWord' -Value '1'

        Creates or sets a DWord value in the specified registry key.

    .EXAMPLE
        Set-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name 'Debugger' -Value $blockedAppDebuggerValue -Type String

        Creates or sets a String value in the specified registry key.

    .EXAMPLE
        Set-ADTRegistryKey -Key 'HKCU\Software\Microsoft\Example' -Name 'Data' -Value (0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x02,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x02,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x00,0x01,0x01,0x01,0x02,0x02,0x02) -Type 'Binary'

        Creates or sets a Binary value in the specified registry key.

    .EXAMPLE
        Set-ADTRegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Example' -Name '(Default)' -Value "Text"

        Creates or sets the default value in the specified registry key.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false)]
        [System.Object]$Value,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Binary', 'DWord', 'ExpandString', 'MultiString', 'None', 'QWord', 'String', 'Unknown')]
        [Microsoft.Win32.RegistryValueKind]$Type = 'String',

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        try
        {
            try
            {
                # If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID.
                $Key = if ($PSBoundParameters.ContainsKey('SID'))
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -Wow6432Node:$Wow6432Node -SID $SID
                }
                else
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -Wow6432Node:$Wow6432Node
                }

                # Create registry key if it doesn't exist.
                if (!(& $Script:CommandTable.'Test-Path' -LiteralPath $Key))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating registry key [$Key]."
                    if (($Key.Split('/').Count - 1) -eq 0)
                    {
                        # No forward slash found in Key. Use New-Item cmdlet to create registry key.
                        $null = & $Script:CommandTable.'New-Item' -Path $Key -ItemType Registry -Force
                    }
                    else
                    {
                        # Forward slash was found in Key. Use REG.exe ADD to create registry key
                        $RegMode = if ([System.Environment]::Is64BitProcess -and !$Wow6432Node)
                        {
                            '/reg:64'
                        }
                        else
                        {
                            '/reg:32'
                        }
                        $null = & "$([System.Environment]::SystemDirectory)\reg.exe" ADD "$($Key.Substring($Key.IndexOf('::') + 2))" /f $RegMode 2>&1
                    }
                }

                if ($Name)
                {
                    if (!(& $Script:CommandTable.'Get-ItemProperty' -LiteralPath $Key -Name $Name -ErrorAction Ignore))
                    {
                        # Set registry value if it doesn't exist.
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Setting registry key value: [$Key] [$Name = $Value]."
                        $null = & $Script:CommandTable.'New-ItemProperty' -LiteralPath $Key -Name $Name -Value $Value -PropertyType $Type
                    }
                    else
                    {
                        # Update registry value if it does exist.
                        if ($Name -eq '(Default)')
                        {
                            # Set Default registry key value with the following workaround, because Set-ItemProperty contains a bug and cannot set Default registry key value.
                            $null = (& $Script:CommandTable.'Get-Item' -LiteralPath $Key).OpenSubKey('', 'ReadWriteSubTree').SetValue($null, $Value)
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Updating registry key value: [$Key] [$Name = $Value]."
                            $null = & $Script:CommandTable.'Set-ItemProperty' -LiteralPath $Key -Name $Name -Value $Value
                        }
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to $(("set registry key [$Key]", "update value [$Value] for registry key [$Key] [$Name]")[!!$Name])."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTServiceStartMode
#
#-----------------------------------------------------------------------------

function Set-ADTServiceStartMode
{
    <#
    .SYNOPSIS
        Set the service startup mode.

    .DESCRIPTION
        Set the service startup mode. This function allows you to configure the startup mode of a specified service. The startup modes available are: Automatic, Automatic (Delayed Start), Manual, Disabled, Boot, and System.

    .PARAMETER Service
        Specify the name of the service.

    .PARAMETER StartMode
        Specify startup mode for the service. Options: Automatic, Automatic (Delayed Start), Manual, Disabled, Boot, System.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Set-ADTServiceStartMode -Service 'wuauserv' -StartMode 'Automatic (Delayed Start)'

        Sets the 'wuauserv' service to start automatically with a delayed start.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!$_.Name)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Service -ProvidedValue $_ -ExceptionMessage 'The specified service does not exist.'))
                }
                return !!$_
            })]
        [System.ServiceProcess.ServiceController]$Service,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Automatic', 'Automatic (Delayed Start)', 'Manual', 'Disabled', 'Boot', 'System')]
        [System.String]$StartMode
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        # Re-write StartMode to suit sc.exe.
        & $Script:CommandTable.'New-Variable' -Name StartMode -Force -Confirm:$false -Value $(switch ($StartMode)
            {
                'Automatic' { 'Auto'; break }
                'Automatic (Delayed Start)' { 'Delayed-Auto'; break }
                'Manual' { 'Demand'; break }
                default { $_; break }
            })
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(($msg = "Setting service [$($Service.Name)] startup mode to [$StartMode]"))."
        try
        {
            try
            {
                # Set the start up mode using sc.exe. Note: we found that the ChangeStartMode method in the Win32_Service WMI class set services to 'Automatic (Delayed Start)' even when you specified 'Automatic' on Win7, Win8, and Win10.
                $scResult = & "$([System.Environment]::SystemDirectory)\sc.exe" config $Service.Name start= $StartMode 2>&1
                if (!$Global:LASTEXITCODE)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Successfully set service [($Service.Name)] startup mode to [$StartMode]."
                    return
                }

                # If we're here, we had a bad exit code.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message ($msg = "$msg failed with exit code [$Global:LASTEXITCODE]: $scResult") -Severity 3
                $naerParams = @{
                    Exception = [System.Runtime.InteropServices.ExternalException]::new($msg, $Global:LASTEXITCODE)
                    Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                    ErrorId = 'ScConfigFailure'
                    TargetObject = $scResult
                    RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                }
                throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Set-ADTShortcut
#
#-----------------------------------------------------------------------------

function Set-ADTShortcut
{
    <#
    .SYNOPSIS
        Modifies a .lnk or .url type shortcut.

    .DESCRIPTION
        Modifies a shortcut - .lnk or .url file, with configurable options. Only specify the parameters that you want to change.

    .PARAMETER Path
        Path to the shortcut to be changed.

    .PARAMETER TargetPath
        Sets target path or URL that the shortcut launches.

    .PARAMETER Arguments
        Sets the arguments used against the target path.

    .PARAMETER IconLocation
        Sets location of the icon used for the shortcut.

    .PARAMETER IconIndex
        Sets the index of the icon. Executables, DLLs, ICO files with multiple icons need the icon index to be specified. This parameter is an Integer. The first index is 0.

    .PARAMETER Description
        Sets the description of the shortcut as can be seen in the shortcut's properties.

    .PARAMETER WorkingDirectory
        Sets working directory to be used for the target path.

    .PARAMETER WindowStyle
        Sets the shortcut's window style to be minimised, maximised, etc.

    .PARAMETER RunAsAdmin
        Sets the shortcut to require elevated permissions to run.

    .PARAMETER HotKey
        Sets the hotkey to launch the shortcut, e.g. "CTRL+SHIFT+F".

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Set-ADTShortcut -Path "$envCommonDesktop\Application.lnk" -TargetPath "$envProgramFiles\Application\application.exe"

        Creates a shortcut on the All Users desktop named 'Application', targeted to '$envProgramFiles\Application\application.exe'.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_) -or (![System.IO.Path]::GetExtension($Path).ToLower().Equals('.lnk') -and ![System.IO.Path]::GetExtension($Path).ToLower().Equals('.url')))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Path -ProvidedValue $_ -ExceptionMessage 'The specified path does not exist or does not have the correct extension.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$TargetPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Arguments,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IconLocation,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$IconIndex,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Description,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Normal', 'Maximized', 'Minimized', 'DontChange')]
        [System.String]$WindowStyle = 'DontChange',

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$RunAsAdmin,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Hotkey
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Changing shortcut [$Path]."
        try
        {
            try
            {
                # Make sure .NET's current directory is synced with PowerShell's.
                [System.IO.Directory]::SetCurrentDirectory((& $Script:CommandTable.'Get-Location' -PSProvider FileSystem).ProviderPath)
                if ($extension -eq '.url')
                {
                    $URLFile = [System.IO.File]::ReadAllLines($Path) | & {
                        process
                        {
                            switch ($_)
                            {
                                { $_.StartsWith('URL=') -and $TargetPath } { "URL=$TargetPath"; break }
                                { $_.StartsWith('IconIndex=') -and ($null -ne $IconIndex) } { "IconIndex=$IconIndex"; break }
                                { $_.StartsWith('IconFile=') -and $IconLocation } { "IconFile=$IconLocation"; break }
                                default { $_; break }
                            }
                        }
                    }
                    [System.IO.File]::WriteAllLines($Path, $URLFile, [System.Text.UTF8Encoding]::new($false))
                }
                else
                {
                    # Open shortcut and set initial properties.
                    $shortcut = [System.Activator]::CreateInstance([System.Type]::GetTypeFromProgID('WScript.Shell')).CreateShortcut($Path)
                    if ($TargetPath)
                    {
                        $shortcut.TargetPath = $TargetPath
                    }
                    if ($Arguments)
                    {
                        $shortcut.Arguments = $Arguments
                    }
                    if ($Description)
                    {
                        $shortcut.Description = $Description
                    }
                    if ($WorkingDirectory)
                    {
                        $shortcut.WorkingDirectory = $WorkingDirectory
                    }
                    if ($Hotkey)
                    {
                        $shortcut.Hotkey = $Hotkey
                    }

                    # Set the WindowStyle based on input.
                    $windowStyleInt = switch ($WindowStyle)
                    {
                        Normal { 1; break }
                        Maximized { 3; break }
                        Minimized { 7; break }
                    }
                    If ($null -ne $windowStyleInt)
                    {
                        $shortcut.WindowStyle = $WindowStyleInt
                    }

                    # Handle icon, starting with retrieval previous value and split the path from the index.
                    $TempIconLocation, $TempIconIndex = $shortcut.IconLocation.Split(',')
                    $IconLocation = if ($IconLocation)
                    {
                        # New icon path was specified. Check whether new icon index was also specified.
                        if ($null -ne $IconIndex)
                        {
                            # Create new icon path from new icon path and new icon index.
                            $IconLocation + ",$IconIndex"
                        }
                        else
                        {
                            # No new icon index was specified as a parameter. We will keep the old one.
                            $IconLocation + ",$TempIconIndex"
                        }
                    }
                    elseif ($null -ne $IconIndex)
                    {
                        # New icon index was specified, but not the icon location. Append it to the icon path from the shortcut.
                        $IconLocation = $TempIconLocation + ",$IconIndex"
                    }
                    if ($IconLocation)
                    {
                        $shortcut.IconLocation = $IconLocation
                    }

                    # Save the changes.
                    $shortcut.Save()

                    # Set shortcut to run program as administrator.
                    if ($PSBoundParameters.ContainsKey('RunAsAdmin'))
                    {
                        $fileBytes = [System.IO.FIle]::ReadAllBytes($Path)
                        $fileBytes[21] = if ($PSBoundParameters.RunAsAdmin)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Setting shortcut to run program as administrator.'
                            $fileBytes[21] -bor 32
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Setting shortcut to not run program as administrator.'
                            $fileBytes[21] -band (-bnot 32)
                        }
                        [System.IO.FIle]::WriteAllBytes($Path, $fileBytes)
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to change the shortcut [$Path]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTBalloonTip
#
#-----------------------------------------------------------------------------

function Show-ADTBalloonTip
{
    <#
    .SYNOPSIS
        Displays a balloon tip notification in the system tray.

    .DESCRIPTION
        Displays a balloon tip notification in the system tray. This function can be used to show notifications to the user with customizable text, title, icon, and display duration.

        For Windows 10 and above, balloon tips automatically get translated by the system into toast notifications.

    .PARAMETER BalloonTipText
        Text of the balloon tip.

    .PARAMETER BalloonTipIcon
        Icon to be used. Options: 'Error', 'Info', 'None', 'Warning'. Default is: Info.

    .PARAMETER BalloonTipTime
        Time in milliseconds to display the balloon tip. Default: 10000.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Show-ADTBalloonTip -BalloonTipText 'Installation Started' -BalloonTipTitle 'Application Name'

        Displays a balloon tip with the text 'Installation Started' and the title 'Application Name'.

    .EXAMPLE
        Show-ADTBalloonTip -BalloonTipIcon 'Info' -BalloonTipText 'Installation Started' -BalloonTipTitle 'Application Name' -BalloonTipTime 1000

        Displays a balloon tip with the info icon, the text 'Installation Started', the title 'Application Name', and a display duration of 1000 milliseconds.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'BalloonTipIcon', Justification = "This parameter is used via the function's PSBoundParameters dictionary, which is not something PSScriptAnalyzer understands. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]$BalloonTipText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.ToolTipIcon]$BalloonTipIcon = 'Info',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$BalloonTipTime = 10000
    )

    dynamicparam
    {
        # Initialize the module first if needed.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('BalloonTipTitle', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'BalloonTipTitle', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Title of the balloon tip.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Set up defaults if not specified.
        if (!$PSBoundParameters.ContainsKey('BalloonTipTitle'))
        {
            $PSBoundParameters.Add('BalloonTipTitle', $adtSession.InstallTitle)
        }
    }

    process
    {
        # Don't allow toast notifications with fluent dialogs unless this function was explicitly requested by the caller.
        if (($adtConfig.UI.DialogStyle -eq 'Fluent') -and ((& $Script:CommandTable.'Get-PSCallStack')[1].Command -match '^(Show|Close)-ADTInstallationProgress$'))
        {
            return
        }

        try
        {
            try
            {
                # Skip balloon if in silent mode, disabled in the config or presentation is detected.
                if (!$adtConfig.UI.BalloonNotifications)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Config Show Balloon Notifications: $($adtConfig.UI.BalloonNotifications)]. BalloonTipText: $BalloonTipText"
                    return
                }
                if ($adtSession -and $adtSession.IsSilent())
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Mode: $($adtSession.DeployMode)]. BalloonTipText: $BalloonTipText"
                    return
                }
                if (& $Script:CommandTable.'Test-ADTUserIsBusy')
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Presentation/Microphone in Use Detected: $true]. BalloonTipText: $BalloonTipText"
                    return
                }

                # Display the balloon tip to the user. As all assets are in memory, there's nothing to dispose.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Displaying balloon tip notification with message [$BalloonTipText]."
                $nabtParams = & $Script:CommandTable.'Get-ADTBoundParametersAndDefaultValues' -Invocation $MyInvocation -Exclude BalloonTipTime
                $nabtParams.Add('Icon', $Script:Dialogs.Classic.Assets.Icon); $nabtParams.Add('Visible', $true)
                ([System.Windows.Forms.NotifyIcon]$nabtParams).ShowBalloonTip($BalloonTipTime)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTDialogBox
#
#-----------------------------------------------------------------------------

function Show-ADTDialogBox
{
    <#
    .SYNOPSIS
        Display a custom dialog box with optional title, buttons, icon, and timeout.

    .DESCRIPTION
        Display a custom dialog box with optional title, buttons, icon, and timeout. The default button is "OK", the default Icon is "None", and the default Timeout is None.

        Show-ADTInstallationPrompt is recommended over this function as it provides more customization and uses consistent branding with the other UI components.

    .PARAMETER Text
        Text in the message dialog box.

    .PARAMETER Buttons
        The button(s) to display on the dialog box.

    .PARAMETER DefaultButton
        The Default button that is selected. Options: First, Second, Third.

    .PARAMETER Icon
        Icon to display on the dialog box. Options: None, Stop, Question, Exclamation, Information.

    .PARAMETER NotTopMost
        Specifies whether the message box shouldn't be a system modal message box that appears in a topmost window.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.String

        Returns the text of the button that was clicked.

    .EXAMPLE
        Show-ADTDialogBox -Title 'Installation Notice' -Text 'Installation will take approximately 30 minutes. Do you wish to proceed?' -Buttons 'OKCancel' -DefaultButton 'Second' -Icon 'Exclamation' -Timeout 600 -Topmost $false

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Enter a message for the dialog box.')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Text,

        [Parameter(Mandatory = $false)]
        [ValidateSet('OK', 'OKCancel', 'AbortRetryIgnore', 'YesNoCancel', 'YesNo', 'RetryCancel', 'CancelTryAgainContinue')]
        [System.String]$Buttons = 'OK',

        [Parameter(Mandatory = $false)]
        [ValidateSet('First', 'Second', 'Third')]
        [System.String]$DefaultButton = 'First',

        [Parameter(Mandatory = $false)]
        [ValidateSet('Exclamation', 'Information', 'None', 'Stop', 'Question')]
        [System.String]$Icon = 'None',

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost
    )

    dynamicparam
    {
        # Initialize the module if there's no session and it hasn't been previously initialized.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('Title', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Title', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Title of the message dialog box.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('Timeout', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Timeout', [System.UInt32], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = $false; HelpMessage = 'Specifies how long, in seconds, to show the message prompt before aborting.' }
                    [System.Management.Automation.ValidateScriptAttribute]::new({
                            if ($_ -gt $adtConfig.UI.DefaultTimeout)
                            {
                                $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Timeout -ProvidedValue $_ -ExceptionMessage 'The installation UI dialog timeout cannot be longer than the timeout specified in the config.psd1 file.'))
                            }
                            return !!$_
                        })
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Set up defaults if not specified.
        $Title = if (!$PSBoundParameters.ContainsKey('Title'))
        {
            $adtSession.InstallTitle
        }
        else
        {
            $PSBoundParameters.Title
        }
        $Timeout = if (!$PSBoundParameters.ContainsKey('Timeout'))
        {
            $adtConfig.UI.DefaultTimeout
        }
        else
        {
            $PSBoundParameters.Timeout
        }
    }

    process
    {
        # Bypass if in silent mode.
        if ($adtSession -and $adtSession.IsNonInteractive())
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Mode: $($adtSession.deployMode)]. Text: $Text"
            return
        }

        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Displaying Dialog Box with message: $Text..."
        try
        {
            try
            {
                $result = switch ((& $Script:CommandTable.'Get-ADTEnvironmentTable').Shell.Popup($Text, $Timeout, $Title, ($Script:Dialogs.Box.Buttons.$Buttons + $Script:Dialogs.Box.Icons.$Icon + $Script:Dialogs.Box.DefaultButtons.$DefaultButton + (4096 * !$NotTopMost))))
                {
                    1 { 'OK'; break }
                    2 { 'Cancel'; break }
                    3 { 'Abort'; break }
                    4 { 'Retry'; break }
                    5 { 'Ignore'; break }
                    6 { 'Yes'; break }
                    7 { 'No'; break }
                    10 { 'Try Again'; break }
                    11 { 'Continue'; break }
                    -1 { 'Timeout'; break }
                    default { 'Unknown'; break }
                }

                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Dialog Box Response: $result"
                return $result
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTHelpConsole
#
#-----------------------------------------------------------------------------

function Show-ADTHelpConsole
{
    <#
    .SYNOPSIS
        Displays a help console for the ADT module.

    .DESCRIPTION
        Displays a help console for the ADT module in a new PowerShell window. The console provides a graphical interface to browse and view detailed help information for all commands exported by the ADT module. The help console includes a list box to select commands and a text box to display the full help content for the selected command.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Show-ADTHelpConsole

        Opens a new PowerShell window displaying the help console for the ADT module.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    # Run this via a new PowerShell window so it doesn't stall the main thread.
    & $Script:CommandTable.'Start-Process' -FilePath (& $Script:CommandTable.'Get-ADTPowerShellProcessPath') -NoNewWindow -ArgumentList "$(if (!(& $Script:CommandTable.'Test-ADTModuleIsReleaseBuild')) { "-ExecutionPolicy Bypass " })-NonInteractive -NoProfile -NoLogo -EncodedCommand $(& $Script:CommandTable.'Out-ADTPowerShellEncodedCommand' -Command "& {$($Script:CommandTable.'Show-ADTHelpConsoleInternal'.ScriptBlock)} -ModuleName '$($Script:PSScriptRoot)\$($MyInvocation.MyCommand.Module.Name).psd1' -Guid $($MyInvocation.MyCommand.Module.Guid) -ModuleVersion $($MyInvocation.MyCommand.Module.Version)")"
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationProgress
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationProgress
{
    <#
    .SYNOPSIS
        Displays a progress dialog in a separate thread with an updateable custom message.

    .DESCRIPTION
        Creates a WPF window in a separate thread to display a marquee style progress ellipse with a custom message that can be updated. The status message supports line breaks.

        The first time this function is called in a script, it will display a balloon tip notification to indicate that the installation has started (provided balloon tips are enabled in the config.psd1 file).

    .PARAMETER WindowLocation
        The location of the progress window. Default: center of the screen.

    .PARAMETER MessageAlignment
        The text alignment to use for the status message. Default: center.

    .PARAMETER NotTopMost
        Specifies whether the progress window shouldn't be topmost. Default: $false.

    .PARAMETER NoRelocation
        Specifies whether to not reposition the window upon updating the message. Default: $false.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Show-ADTInstallationProgress

        Uses the default status message from the strings.psd1 file.

    .EXAMPLE
        Show-ADTInstallationProgress -StatusMessage 'Installation in Progress...'

        Displays a progress dialog with the status message 'Installation in Progress...'.

    .EXAMPLE
        Show-ADTInstallationProgress -StatusMessage "Installation in Progress...`nThe installation may take 20 minutes to complete."

        Displays a progress dialog with a multiline status message.

    .EXAMPLE
        Show-ADTInstallationProgress -StatusMessage 'Installation in Progress...' -WindowLocation 'BottomRight' -NotTopMost

        Displays a progress dialog with the status message 'Installation in Progress...', positioned at the bottom right of the screen, and not set as topmost.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Default', 'TopLeft', 'Top', 'TopRight', 'TopCenter', 'BottomLeft', 'Bottom', 'BottomRight')]
        [System.String]$WindowLocation = 'Default',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.TextAlignment]$MessageAlignment,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoRelocation
    )

    dynamicparam
    {
        # Initialize the module first if needed.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('WindowTitle', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'WindowTitle', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'The title of the window to be displayed. The default is the derived value from "$($adtSession.InstallTitle)".' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('WindowSubtitle', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'WindowSubtitle', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession -and ($adtConfig.UI.DialogStyle -eq 'Fluent'); HelpMessage = 'The subtitle of the window to be displayed with a fluent progress window. The default is the derived value from "$($adtSession.DeploymentType)".' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('StatusMessage', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'StatusMessage', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'The status message to be displayed. The default status message is taken from the config.psd1 file.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('StatusMessageDetail', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'StatusMessageDetail', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession -and ($adtConfig.UI.DialogStyle -eq 'Fluent'); HelpMessage = 'The status message detail to be displayed with a fluent progress window. The default status message is taken from the config.psd1 file.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'
        $errRecord = $null

        # Set up defaults if not specified.
        if (!$PSBoundParameters.ContainsKey('WindowTitle'))
        {
            $PSBoundParameters.Add('WindowTitle', $adtSession.InstallTitle)
        }
        if (!$PSBoundParameters.ContainsKey('WindowSubtitle'))
        {
            $PSBoundParameters.Add('WindowSubtitle', [System.String]::Format($adtStrings.WelcomePrompt.Fluent.Subtitle, $adtSession.DeploymentType))
        }
        if (!$PSBoundParameters.ContainsKey('StatusMessage'))
        {
            $PSBoundParameters.Add('StatusMessage', $adtStrings.Progress."Message$($adtSession.DeploymentType)")
        }
        if (!$PSBoundParameters.ContainsKey('StatusMessageDetail') -and ($adtConfig.UI.DialogStyle -eq 'Fluent'))
        {
            $PSBoundParameters.Add('StatusMessageDetail', $adtStrings.Progress."Message$($adtSession.DeploymentType)Detail")
        }
    }

    process
    {
        # Determine if progress window is open before proceeding.
        $progressOpen = & $Script:CommandTable.'Test-ADTInstallationProgressRunning'

        # Return early in silent mode.
        if ($adtSession)
        {
            if ($adtSession.IsSilent())
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Mode: $($adtSession.DeployMode)]. Status message: $($PSBoundParameters.StatusMessage)"
                return
            }

            # Notify user that the software installation has started.
            if (!$progressOpen)
            {
                try
                {
                    & $Script:CommandTable.'Show-ADTBalloonTip' -BalloonTipIcon Info -BalloonTipText "$($adtSession.GetDeploymentTypeName()) $($adtStrings.BalloonText.Start)"
                }
                catch
                {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            }
        }

        # Call the underlying function to open the progress window.
        try
        {
            try
            {
                # Perform the dialog action.
                if (!$progressOpen)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Creating the progress dialog in a separate thread with message: [$($PSBoundParameters.StatusMessage)]."
                }
                else
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Updating the progress dialog with message: [$($PSBoundParameters.StatusMessage)]."
                }
                & $Script:CommandTable."$($MyInvocation.MyCommand.Name)$($adtConfig.UI.DialogStyle)" @PSBoundParameters

                # Add a callback to close it if we've opened for the first time.
                if (!(& $Script:CommandTable.'Test-ADTInstallationProgressRunning').Equals($progressOpen))
                {
                    & $Script:CommandTable.'Add-ADTSessionFinishingCallback' -Callback $Script:CommandTable.'Close-ADTInstallationProgress'
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord ($errRecord = $_)
        }
        finally
        {
            if ($errRecord)
            {
                & $Script:CommandTable.'Close-ADTInstallationProgress'
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationPrompt
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationPrompt
{
    <#
    .SYNOPSIS
        Displays a custom installation prompt with the toolkit branding and optional buttons.

    .DESCRIPTION
        Displays a custom installation prompt with the toolkit branding and optional buttons. Any combination of Left, Middle, or Right buttons can be displayed. The return value of the button clicked by the user is the button text specified. The prompt can also display a system icon and be configured to persist, minimize other windows, or timeout after a specified period.

    .PARAMETER Message
        The message text to be displayed on the prompt.

    .PARAMETER MessageAlignment
        Alignment of the message text. Options: Left, Center, Right. Default: Center.

    .PARAMETER ButtonLeftText
        Show a button on the left of the prompt with the specified text.

    .PARAMETER ButtonRightText
        Show a button on the right of the prompt with the specified text.

    .PARAMETER ButtonMiddleText
        Show a button in the middle of the prompt with the specified text.

    .PARAMETER Icon
        Show a system icon in the prompt. Options: Application, Asterisk, Error, Exclamation, Hand, Information, None, Question, Shield, Warning, WinLogo. Default: None.

    .PARAMETER NoWait
        Presents the dialog in a separate, independent thread so that the main process isn't stalled waiting for a response.

    .PARAMETER PersistPrompt
        Specify whether to make the prompt persist in the center of the screen every couple of seconds, specified in the AppDeployToolkitConfig.xml. The user will have no option but to respond to the prompt - resistance is futile!

    .PARAMETER MinimizeWindows
        Specifies whether to minimize other windows when displaying prompt.

    .PARAMETER NoExitOnTimeout
        Specifies whether to not exit the script if the UI times out.

    .PARAMETER NotTopMost
        Specifies whether the prompt shouldn't be topmost, above all other windows.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Show-ADTInstallationPrompt -Message 'Do you want to proceed with the installation?' -ButtonRightText 'Yes' -ButtonLeftText 'No'

    .EXAMPLE
        Show-ADTInstallationPrompt -Title 'Funny Prompt' -Message 'How are you feeling today?' -ButtonRightText 'Good' -ButtonLeftText 'Bad' -ButtonMiddleText 'Indifferent'

    .EXAMPLE
        Show-ADTInstallationPrompt -Message 'You can customize text to appear at the end of an install, or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Left', 'Center', 'Right')]
        [System.String]$MessageAlignment = 'Center',

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonRightText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonLeftText,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ButtonMiddleText,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Application', 'Asterisk', 'Error', 'Exclamation', 'Hand', 'Information', 'Question', 'Shield', 'Warning', 'WinLogo')]
        [System.String]$Icon,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PersistPrompt,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$MinimizeWindows,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoExitOnTimeout,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost
    )

    dynamicparam
    {
        # Initialize variables.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('Title', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Title', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Title of the prompt. Default: the application installation name.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('Subtitle', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Subtitle', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Subtitle of the prompt. Default: the application deployment type.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('Timeout', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Timeout', [System.UInt32], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = $false; HelpMessage = 'Specifies how long, in seconds, to show the message prompt before aborting.' }
                    [System.Management.Automation.ValidateScriptAttribute]::new({
                            if ($_ -gt $adtConfig.UI.DefaultTimeout)
                            {
                                $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Timeout -ProvidedValue $_ -ExceptionMessage 'The installation UI dialog timeout cannot be longer than the timeout specified in the config.psd1 file.'))
                            }
                            return !!$_
                        })
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Throw a terminating error if at least one button isn't specified.
        if (!($PSBoundParameters.Keys -match '^Button'))
        {
            $naerParams = @{
                Exception = [System.ArgumentException]::new('At least one button must be specified when calling this function.')
                Category = [System.Management.Automation.ErrorCategory]::InvalidArgument
                ErrorId = 'MandatoryParameterMissing'
                TargetObject = $PSBoundParameters
                RecommendedAction = "Please review the supplied parameters used against $($MyInvocation.MyCommand.Name) and try again."
            }
            $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
        }

        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Set up defaults if not specified.
        if (!$PSBoundParameters.ContainsKey('Title'))
        {
            $PSBoundParameters.Add('Title', $adtSession.InstallTitle)
        }
        if (!$PSBoundParameters.ContainsKey('Subtitle'))
        {
            $PSBoundParameters.Add('Subtitle', [System.String]::Format((& $Script:CommandTable.'Get-ADTStringTable').WelcomePrompt.Fluent.Subtitle, $adtSession.DeploymentType))
        }
        if (!$PSBoundParameters.ContainsKey('Timeout'))
        {
            $PSBoundParameters.Add('Timeout', $adtConfig.UI.DefaultTimeout)
        }
    }

    process
    {
        try
        {
            try
            {
                # Bypass if in non-interactive mode.
                if ($adtSession -and $adtSession.IsNonInteractive())
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing $($MyInvocation.MyCommand.Name) [Mode: $($adtSession.DeployMode)]. Message: $Message"
                    return
                }

                # Resolve the bound parameters to a string.
                $paramsString = [PSADT.Shared.Utility]::ConvertDictToPowerShellArgs($PSBoundParameters)

                # If the NoWait parameter is specified, launch a new PowerShell session to show the prompt asynchronously.
                if ($NoWait)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Displaying custom installation prompt asynchronously with the parameters: [$($paramsString.Replace("''", "'"))]."
                    & $Script:CommandTable.'Start-Process' -FilePath (& $Script:CommandTable.'Get-ADTPowerShellProcessPath') -ArgumentList "$(if (!(& $Script:CommandTable.'Test-ADTModuleIsReleaseBuild')) { "-ExecutionPolicy Bypass " })-NonInteractive -NoProfile -NoLogo -WindowStyle Hidden -Command & (Import-Module -FullyQualifiedName @{ ModuleName = '$("$($Script:PSScriptRoot)\$($MyInvocation.MyCommand.Module.Name).psd1".Replace("'", "''"))'; Guid = '$($MyInvocation.MyCommand.Module.Guid)'; ModuleVersion = '$($MyInvocation.MyCommand.Module.Version)' } -PassThru) { & `$CommandTable.'Initialize-ADTModule' -ScriptDirectory '$($Script:ADT.Directories.Script.Replace("'", "''"))'; `$null = & `$CommandTable.'$($MyInvocation.MyCommand.Name)$($adtConfig.UI.DialogStyle)' $($paramsString.Replace('"', '\"')) }" -WindowStyle Hidden -ErrorAction Ignore
                    return
                }

                # Close the Installation Progress dialog if running.
                if ($adtSession)
                {
                    & $Script:CommandTable.'Close-ADTInstallationProgress'
                }

                # Call the underlying function to open the message prompt.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Displaying custom installation prompt with the parameters: [$($paramsString.Replace("''", "'"))]."
                return & $Script:CommandTable."$($MyInvocation.MyCommand.Name)$($adtConfig.UI.DialogStyle)" @PSBoundParameters
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationRestartPrompt
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationRestartPrompt
{
    <#
    .SYNOPSIS
        Displays a restart prompt with a countdown to a forced restart.

    .DESCRIPTION
        Displays a restart prompt with a countdown to a forced restart. The prompt can be customized with a title, countdown duration, and whether it should be topmost. It also supports silent mode where the restart can be triggered without user interaction.

    .PARAMETER CountdownSeconds
        Specifies the number of seconds to display the restart prompt. Default: 60

    .PARAMETER CountdownNoHideSeconds
        Specifies the number of seconds to display the restart prompt without allowing the window to be hidden. Default: 30

    .PARAMETER SilentCountdownSeconds
        Specifies number of seconds to countdown for the restart when the toolkit is running in silent mode and NoSilentRestart is $false. Default: 5

    .PARAMETER SilentRestart
        Specifies whether the restart should be triggered when Deploy mode is silent or very silent.

    .PARAMETER NoCountdown
        Specifies whether the user should receive a prompt to immediately restart their workstation.

    .PARAMETER NotTopMost
        Specifies whether the prompt shouldn't be topmost, above all other windows.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Show-ADTInstallationRestartPrompt -NoCountdown

        Displays a restart prompt without a countdown.

    .EXAMPLE
        Show-ADTInstallationRestartPrompt -Countdownseconds 300

        Displays a restart prompt with a 300-second countdown.

    .EXAMPLE
        Show-ADTInstallationRestartPrompt -CountdownSeconds 600 -CountdownNoHideSeconds 60

        Displays a restart prompt with a 600-second countdown and triggers a silent restart with a 60-second countdown in silent mode.

    .NOTES
        Be mindful of the countdown you specify for the reboot as code directly after this function might NOT be able to execute - that includes logging.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$CountdownSeconds = 60,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$CountdownNoHideSeconds = 30,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$SilentCountdownSeconds = 5,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SilentRestart,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoCountdown,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NotTopMost
    )

    dynamicparam
    {
        # Initialize variables.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('Title', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Title', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Title of the prompt. Default: the application installation name.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('Subtitle', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Subtitle', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = 'Subtitle of the prompt. Default: the application deployment type.' }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

        # Set up defaults if not specified.
        if (!$PSBoundParameters.ContainsKey('Title'))
        {
            $PSBoundParameters.Add('Title', $adtSession.InstallTitle)
        }
        if (!$PSBoundParameters.ContainsKey('Subtitle'))
        {
            $PSBoundParameters.Add('Subtitle', [System.String]::Format($adtStrings.WelcomePrompt.Fluent.Subtitle, $adtSession.DeploymentType))
        }
        if (!$PSBoundParameters.ContainsKey('CountdownSeconds'))
        {
            $PSBoundParameters.Add('CountdownSeconds', $CountdownSeconds)
        }
        if (!$PSBoundParameters.ContainsKey('CountdownNoHideSeconds'))
        {
            $PSBoundParameters.Add('CountdownNoHideSeconds', $CountdownNoHideSeconds)
        }
    }

    process
    {
        try
        {
            try
            {
                # If in non-interactive mode.
                if ($adtSession -and $adtSession.IsSilent())
                {
                    if ($SilentRestart)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Triggering restart silently, because the deploy mode is set to [$($adtSession.DeployMode)] and [NoSilentRestart] is disabled. Timeout is set to [$SilentCountdownSeconds] seconds."
                        & $Script:CommandTable.'Start-Process' -FilePath (& $Script:CommandTable.'Get-ADTPowerShellProcessPath') -ArgumentList "-NonInteractive -NoProfile -NoLogo -WindowStyle Hidden -Command Start-Sleep -Seconds $SilentCountdownSeconds; Restart-Computer -Force" -WindowStyle Hidden -ErrorAction Ignore
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Skipping restart, because the deploy mode is set to [$($adtSession.DeployMode)] and [SilentRestart] is false."
                    }
                    return
                }

                # Check if we are already displaying a restart prompt.
                if (& $Script:CommandTable.'Get-Process' | & { process { if ($_.MainWindowTitle -match $adtStrings.RestartPrompt.Title) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "$($MyInvocation.MyCommand.Name) was invoked, but an existing restart prompt was detected. Cancelling restart prompt." -Severity 2
                    return
                }

                # If the script has been dot-source invoked by the deploy app script, display the restart prompt asynchronously.
                if ($adtSession)
                {
                    if ($NoCountdown)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoking $($MyInvocation.MyCommand.Name) asynchronously with no countdown..."
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoking $($MyInvocation.MyCommand.Name) asynchronously with a [$CountdownSeconds] second countdown..."
                    }

                    # Start another powershell instance silently with function parameters from this function.
                    & $Script:CommandTable.'Start-Process' -FilePath (& $Script:CommandTable.'Get-ADTPowerShellProcessPath') -ArgumentList "$(if (!(& $Script:CommandTable.'Test-ADTModuleIsReleaseBuild')) { "-ExecutionPolicy Bypass " })-NonInteractive -NoProfile -NoLogo -WindowStyle Hidden -Command & (Import-Module -FullyQualifiedName @{ ModuleName = '$("$($Script:PSScriptRoot)\$($MyInvocation.MyCommand.Module.Name).psd1".Replace("'", "''"))'; Guid = '$($MyInvocation.MyCommand.Module.Guid)'; ModuleVersion = '$($MyInvocation.MyCommand.Module.Version)' } -PassThru) { & `$CommandTable.'Initialize-ADTModule' -ScriptDirectory '$($Script:ADT.Directories.Script.Replace("'", "''"))'; `$null = & `$CommandTable.'$($MyInvocation.MyCommand.Name)$($adtConfig.UI.DialogStyle)' $([PSADT.Shared.Utility]::ConvertDictToPowerShellArgs($PSBoundParameters, ('SilentRestart', 'SilentCountdownSeconds')).Replace('"', '\"')) }" -WindowStyle Hidden -ErrorAction Ignore
                    return
                }

                # Call the underlying function to open the restart prompt.
                return & $Script:CommandTable."$($MyInvocation.MyCommand.Name)$($adtConfig.UI.DialogStyle)" @PSBoundParameters
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Show-ADTInstallationWelcome
#
#-----------------------------------------------------------------------------

function Show-ADTInstallationWelcome
{
    <#
    .SYNOPSIS
        Show a welcome dialog prompting the user with information about the installation and actions to be performed before the installation can begin.

    .DESCRIPTION
        The following prompts can be included in the welcome dialog:
            a) Close the specified running applications, or optionally close the applications without showing a prompt (using the -Silent switch).
            b) Defer the installation a certain number of times, for a certain number of days or until a deadline is reached.
            c) Countdown until applications are automatically closed.
            d) Prevent users from launching the specified applications while the installation is in progress.

    .PARAMETER CloseProcesses
        Name of the process to stop (do not include the .exe). Specify multiple processes separated by a comma. Specify custom descriptions like this: @{ Name = 'winword'; Description = 'Microsoft Office Word'},@{ Name = 'excel'; Description = 'Microsoft Office Excel'}

    .PARAMETER Silent
        Stop processes without prompting the user.

    .PARAMETER CloseProcessesCountdown
        Option to provide a countdown in seconds until the specified applications are automatically closed. This only takes effect if deferral is not allowed or has expired.

    .PARAMETER ForceCloseProcessesCountdown
        Option to provide a countdown in seconds until the specified applications are automatically closed regardless of whether deferral is allowed.

    .PARAMETER PromptToSave
        Specify whether to prompt to save working documents when the user chooses to close applications by selecting the "Close Programs" button. Option does not work in SYSTEM context unless toolkit launched with "psexec.exe -s -i" to run it as an interactive process under the SYSTEM account.

    .PARAMETER PersistPrompt
        Specify whether to make the Show-ADTInstallationWelcome prompt persist in the center of the screen every couple of seconds, specified in the AppDeployToolkitConfig.xml. The user will have no option but to respond to the prompt. This only takes effect if deferral is not allowed or has expired.

    .PARAMETER BlockExecution
        Option to prevent the user from launching processes/applications, specified in -CloseProcesses, during the installation.

    .PARAMETER AllowDefer
        Enables an optional defer button to allow the user to defer the installation.

    .PARAMETER AllowDeferCloseProcesses
        Enables an optional defer button to allow the user to defer the installation only if there are running applications that need to be closed. This parameter automatically enables -AllowDefer

    .PARAMETER DeferTimes
        Specify the number of times the installation can be deferred.

    .PARAMETER DeferDays
        Specify the number of days since first run that the installation can be deferred. This is converted to a deadline.

    .PARAMETER DeferDeadline
        Specify the deadline date until which the installation can be deferred.

        Specify the date in the local culture if the script is intended for that same culture.

        If the script is intended to run on EN-US machines, specify the date in the format: "08/25/2013" or "08-25-2013" or "08-25-2013 18:00:00"

        If the script is intended for multiple cultures, specify the date in the universal sortable date/time format: "2013-08-22 11:51:52Z"

        The deadline date will be displayed to the user in the format of their culture.

    .PARAMETER CheckDiskSpace
        Specify whether to check if there is enough disk space for the installation to proceed.

        If this parameter is specified without the RequiredDiskSpace parameter, the required disk space is calculated automatically based on the size of the script source and associated files.

    .PARAMETER RequiredDiskSpace
        Specify required disk space in MB, used in combination with CheckDiskSpace.

    .PARAMETER NoMinimizeWindows
        Specifies whether to minimize other windows when displaying prompt. Default: $false.

    .PARAMETER TopMost
        Specifies whether the windows is the topmost window. Default: $true.

    .PARAMETER ForceCountdown
        Specify a countdown to display before automatically proceeding with the installation when a deferral is enabled.

    .PARAMETER CustomText
        Specify whether to display a custom message specified in the string.psd1 file. Custom message must be populated for each language section in the string.psd1 file.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any output.

    .EXAMPLE
        Show-ADTInstallationWelcome -CloseProcesses iexplore, winword, excel

        Prompt the user to close Internet Explorer, Word and Excel.

    .EXAMPLE
        Show-ADTInstallationWelcome -CloseProcesses @{ Name = 'winword' }, @{ Name = 'excel' } -Silent

        Close Word and Excel without prompting the user.

    .EXAMPLE
        Show-ADTInstallationWelcome -CloseProcesses @{ Name = 'winword' }, @{ Name = 'excel' } -BlockExecution

        Close Word and Excel and prevent the user from launching the applications while the installation is in progress.

    .EXAMPLE
        Show-ADTInstallationWelcome -CloseProcesses @{ Name = 'winword'; Description = 'Microsoft Office Word' }, @{ Name = 'excel'; Description = 'Microsoft Office Excel' } -CloseProcessesCountdown 600

        Prompt the user to close Word and Excel, with customized descriptions for the applications and automatically close the applications after 10 minutes.

    .EXAMPLE
        Show-ADTInstallationWelcome -CloseProcesses @{ Name = 'winword' }, @{ Name = 'msaccess' }, @{ Name = 'excel' } -PersistPrompt

        Prompt the user to close Word, MSAccess and Excel. By using the PersistPrompt switch, the dialog will return to the center of the screen every couple of seconds, specified in the AppDeployToolkitConfig.xml, so the user cannot ignore it by dragging it aside.

    .EXAMPLE
        Show-ADTInstallationWelcome -AllowDefer -DeferDeadline '25/08/2013'

        Allow the user to defer the installation until the deadline is reached.

    .EXAMPLE
        Show-ADTInstallationWelcome -CloseProcesses @{ Name = 'winword' }, @{ Name = 'excel' } -BlockExecution -AllowDefer -DeferTimes 10 -DeferDeadline '25/08/2013' -CloseProcessesCountdown 600

        Close Word and Excel and prevent the user from launching the applications while the installation is in progress.

        Allow the user to defer the installation a maximum of 10 times or until the deadline is reached, whichever happens first.

        When deferral expires, prompt the user to close the applications and automatically close them after 10 minutes.

    .NOTES
        An active ADT session is NOT required to use this function.

        The process descriptions are retrieved via Get-Process, with a fall back on the process name if no description is available. Alternatively, you can specify the description yourself with a '=' symbol - see examples.

        The dialog box will timeout after the timeout specified in the config.psd1 file (default 55 minutes) to prevent Intune/SCCM installations from timing out and returning a failure code. When the dialog times out, the script will exit and return a 1618 code (SCCM fast retry code).

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(DefaultParameterSetName = 'None')]
    param
    (
        [Parameter(Mandatory = $false, HelpMessage = 'Specify process names and an optional process description, e.g. @{ Name = "winword"; Description = "Microsoft Word"}')]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.ProcessObject[]]$CloseProcesses,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify whether to prompt user or force close the applications.')]
        [System.Management.Automation.SwitchParameter]$Silent,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify a countdown to display before automatically closing applications where deferral is not allowed or has expired.')]
        [ValidateNotNullOrEmpty()]
        [System.Double]$CloseProcessesCountdown,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify a countdown to display before automatically closing applications whether or not deferral is allowed.')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$ForceCloseProcessesCountdown,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify whether to prompt to save working documents when the user chooses to close applications by selecting the "Close Programs" button.')]
        [System.Management.Automation.SwitchParameter]$PromptToSave,

        [Parameter(Mandatory = $false, HelpMessage = ' Specify whether to make the prompt persist in the center of the screen every couple of seconds, specified in the AppDeployToolkitConfig.xml.')]
        [System.Management.Automation.SwitchParameter]$PersistPrompt,

        [Parameter(Mandatory = $false, HelpMessage = ' Specify whether to block execution of the processes during installation.')]
        [System.Management.Automation.SwitchParameter]$BlockExecution,

        [Parameter(Mandatory = $false, HelpMessage = ' Specify whether to enable the optional defer button on the dialog box.')]
        [System.Management.Automation.SwitchParameter]$AllowDefer,

        [Parameter(Mandatory = $false, HelpMessage = ' Specify whether to enable the optional defer button on the dialog box only if an app needs to be closed.')]
        [System.Management.Automation.SwitchParameter]$AllowDeferCloseProcesses,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify the number of times the deferral is allowed.')]
        [ValidateNotNullOrEmpty()]
        [System.Int32]$DeferTimes,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify the number of days since first run that the deferral is allowed.')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$DeferDays,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify the deadline (in format dd/mm/yyyy) for which deferral will expire as an option.')]
        [ValidateNotNullOrEmpty()]
        [System.String]$DeferDeadline,

        [Parameter(Mandatory = $true, HelpMessage = 'Specify whether to check if there is enough disk space for the installation to proceed. If this parameter is specified without the RequiredDiskSpace parameter, the required disk space is calculated automatically based on the size of the script source and associated files.', ParameterSetName = 'CheckDiskSpace')]
        [System.Management.Automation.SwitchParameter]$CheckDiskSpace,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify required disk space in MB, used in combination with $CheckDiskSpace.', ParameterSetName = 'CheckDiskSpace')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$RequiredDiskSpace,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify whether to minimize other windows when displaying prompt.')]
        [System.Management.Automation.SwitchParameter]$NoMinimizeWindows,

        [Parameter(Mandatory = $false, HelpMessage = 'Specifies whether the window is the topmost window.')]
        [System.Management.Automation.SwitchParameter]$NotTopMost,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify a countdown to display before automatically proceeding with the installation when a deferral is enabled.')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$ForceCountdown,

        [Parameter(Mandatory = $false, HelpMessage = 'Specify whether to display a custom message specified in the string.psd1 file. Custom message must be populated for each language section in the string.psd1 file.')]
        [System.Management.Automation.SwitchParameter]$CustomText
    )

    dynamicparam
    {
        # Initialize variables.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        $adtStrings = & $Script:CommandTable.'Get-ADTStringTable'
        $adtConfig = & $Script:CommandTable.'Get-ADTConfig'

        # Define parameter dictionary for returning at the end.
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        # Add in parameters we need as mandatory when there's no active ADTSession.
        $paramDictionary.Add('Title', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Title', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = "Title of the prompt. Default: the application installation name." }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('Subtitle', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'Subtitle', [System.String], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession -and ($adtConfig.UI.DialogStyle -eq 'Fluent'); HelpMessage = "Subtitle of the prompt. Default: the application deployment type." }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))
        $paramDictionary.Add('DeploymentType', [System.Management.Automation.RuntimeDefinedParameter]::new(
                'DeploymentType', [PSADT.Module.DeploymentType], $(
                    [System.Management.Automation.ParameterAttribute]@{ Mandatory = !$adtSession; HelpMessage = "The deployment type. Default: the session's DeploymentType value." }
                    [System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new()
                )
            ))

        # Return the populated dictionary.
        return $paramDictionary
    }

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $adtEnv = & $Script:CommandTable.'Get-ADTEnvironmentTable'

        # Set up defaults if not specified.
        if (!$PSBoundParameters.ContainsKey('DeploymentType'))
        {
            $PSBoundParameters.Add('DeploymentType', $adtSession.DeploymentType)
        }
        if (!$PSBoundParameters.ContainsKey('Title'))
        {
            $PSBoundParameters.Add('Title', $adtSession.InstallTitle)
        }
        if (!$PSBoundParameters.ContainsKey('Subtitle'))
        {
            $PSBoundParameters.Add('Subtitle', [System.String]::Format($adtStrings.WelcomePrompt.Fluent.Subtitle, $PSBoundParameters.DeploymentType))
        }

        # Instantiate new object to hold all data needed within this call.
        $welcomeState = [PSADT.Types.WelcomeState]::new()
        $deferDeadlineUniversal = $null
        $promptResult = $null
    }

    process
    {
        try
        {
            try
            {
                # If running in NonInteractive mode, force the processes to close silently.
                if ($adtSession -and $adtSession.IsNonInteractive())
                {
                    $Silent = $true
                }

                # If using Zero-Config MSI Deployment, append any executables found in the MSI to the CloseProcesses list
                if ($adtSession -and ($msiExecutables = $adtSession.GetDefaultMsiExecutablesList()))
                {
                    $CloseProcesses = $(if ($CloseProcesses) { $CloseProcesses }; $msiExecutables)
                }

                # Check disk space requirements if specified
                if ($adtSession -and $CheckDiskSpace)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Evaluating disk space requirements.'
                    if (!$RequiredDiskSpace)
                    {
                        try
                        {
                            # Determine the size of the Files folder
                            $fso = & $Script:CommandTable.'New-Object' -ComObject Scripting.FileSystemObject
                            $RequiredDiskSpace = [System.Math]::Round($fso.GetFolder($adtSession.ScriptDirectory).Size / 1MB)
                        }
                        catch
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to calculate disk space requirement from source files.`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                        }
                        finally
                        {
                            $null = try
                            {
                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fso)
                            }
                            catch
                            {
                                $null
                            }
                        }
                    }
                    if (($freeDiskSpace = & $Script:CommandTable.'Get-ADTFreeDiskSpace') -lt $RequiredDiskSpace)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to meet minimum disk space requirement. Space Required [$RequiredDiskSpace MB], Space Available [$freeDiskSpace MB]." -Severity 3
                        if (!$Silent)
                        {
                            & $Script:CommandTable.'Show-ADTInstallationPrompt' -Message ((& $Script:CommandTable.'Get-ADTStringTable').DiskSpace.Message -f $PSBoundParameters.Title, $RequiredDiskSpace, $freeDiskSpace) -ButtonRightText OK -Icon Error
                        }
                        & $Script:CommandTable.'Close-ADTSession' -ExitCode $adtConfig.UI.DefaultExitCode
                    }
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Successfully passed minimum disk space requirement check.'
                }

                # Check Deferral history and calculate remaining deferrals.
                if ($AllowDefer -or $AllowDeferCloseProcesses)
                {
                    # Set $AllowDefer to true if $AllowDeferCloseProcesses is true.
                    $AllowDefer = $true

                    # Get the deferral history from the registry.
                    $deferHistory = if ($adtSession) { & $Script:CommandTable.'Get-ADTDeferHistory' }
                    $deferHistoryTimes = $deferHistory | & $Script:CommandTable.'Select-Object' -ExpandProperty DeferTimesRemaining -ErrorAction Ignore
                    $deferHistoryDeadline = $deferHistory | & $Script:CommandTable.'Select-Object' -ExpandProperty DeferDeadline -ErrorAction Ignore

                    # Reset switches.
                    $checkDeferDays = $DeferDays -ne 0
                    $checkDeferDeadline = !!$DeferDeadline

                    if ($DeferTimes -ne 0)
                    {
                        $DeferTimes = if ($deferHistoryTimes -ge 0)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Defer history shows [$($deferHistory.DeferTimesRemaining)] deferrals remaining."
                            $deferHistory.DeferTimesRemaining - 1
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "The user has [$DeferTimes] deferrals remaining."
                            $DeferTimes - 1
                        }

                        if ($DeferTimes -lt 0)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Deferral has expired.'
                            $AllowDefer = $false
                        }
                    }

                    if ($checkDeferDays -and $AllowDefer)
                    {
                        $deferDeadlineUniversal = if ($deferHistoryDeadline)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Defer history shows a deadline date of [$deferHistoryDeadline]."
                            & $Script:CommandTable.'Get-ADTUniversalDate' -DateTime $deferHistoryDeadline
                        }
                        else
                        {
                            & $Script:CommandTable.'Get-ADTUniversalDate' -DateTime ([System.DateTime]::Now.AddDays($DeferDays).ToString([System.Globalization.DateTimeFormatInfo]::CurrentInfo.UniversalSortableDateTimePattern))
                        }
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "The user has until [$deferDeadlineUniversal] before deferral expires."

                        if ((& $Script:CommandTable.'Get-ADTUniversalDate') -gt $deferDeadlineUniversal)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Deferral has expired.'
                            $AllowDefer = $false
                        }
                    }

                    if ($checkDeferDeadline -and $AllowDefer)
                    {
                        # Validate date.
                        try
                        {
                            $deferDeadlineUniversal = & $Script:CommandTable.'Get-ADTUniversalDate' -DateTime $DeferDeadline
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "The user has until [$deferDeadlineUniversal] remaining."

                            if ((& $Script:CommandTable.'Get-ADTUniversalDate') -gt $deferDeadlineUniversal)
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Deferral has expired.'
                                $AllowDefer = $false
                            }
                        }
                        catch
                        {
                            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Date is not in the correct format for the current culture. Type the date in the current locale format, such as 20/08/2014 (Europe) or 08/20/2014 (United States). If the script is intended for multiple cultures, specify the date in the universal sortable date/time format, e.g. '2013-08-22 11:51:52Z'."
                        }
                    }
                }

                if (($DeferTimes -lt 0) -and !$deferDeadlineUniversal)
                {
                    $AllowDefer = $false
                }

                # Prompt the user to close running applications and optionally defer if enabled.
                if (!$Silent -and (!$adtSession -or !$adtSession.IsSilent()))
                {
                    # Keep the same variable for countdown to simplify the code.
                    if ($ForceCloseProcessesCountdown -gt 0)
                    {
                        $CloseProcessesCountdown = $ForceCloseProcessesCountdown
                    }
                    elseif ($ForceCountdown -gt 0)
                    {
                        $CloseProcessesCountdown = $ForceCountdown
                    }
                    $welcomeState.CloseProcessesCountdown = $CloseProcessesCountdown

                    while (($welcomeState.RunningProcesses = & $Script:CommandTable.'Get-ADTRunningProcesses' -ProcessObjects $CloseProcesses) -or (($promptResult -ne 'Defer') -and ($promptResult -ne 'Close')))
                    {
                        # Get all unique running process descriptions.
                        $welcomeState.RunningProcessDescriptions = $welcomeState.RunningProcesses | & $Script:CommandTable.'Select-Object' -ExpandProperty ProcessDescription | & $Script:CommandTable.'Sort-Object' -Unique

                        # Define parameters for welcome prompt.
                        $promptParams = @{
                            WelcomeState = $welcomeState
                            Title = $PSBoundParameters.Title
                            Subtitle = $PSBoundParameters.Subtitle
                            DeploymentType = $PSBoundParameters.DeploymentType
                            CloseProcessesCountdown = $welcomeState.CloseProcessesCountdown
                            ForceCloseProcessesCountdown = !!$ForceCloseProcessesCountdown
                            ForceCountdown = !!$ForceCountdown
                            PersistPrompt = $PersistPrompt
                            NoMinimizeWindows = $NoMinimizeWindows
                            CustomText = $CustomText
                            NotTopMost = $NotTopMost
                        }
                        if ($CloseProcesses) { $promptParams.Add('ProcessObjects', $CloseProcesses) }

                        # Check if we need to prompt the user to defer, to defer and close apps, or not to prompt them at all
                        if ($AllowDefer)
                        {
                            # If there is deferral and closing apps is allowed but there are no apps to be closed, break the while loop.
                            if ($AllowDeferCloseProcesses -and !$welcomeState.RunningProcessDescriptions)
                            {
                                break
                            }
                            elseif (($promptResult -ne 'Close') -or ($welcomeState.RunningProcessDescriptions -and ($promptResult -ne 'Continue')))
                            {
                                # Otherwise, as long as the user has not selected to close the apps or the processes are still running and the user has not selected to continue, prompt user to close running processes with deferral.
                                $deferParams = @{ AllowDefer = $true; DeferTimes = $DeferTimes }; if ($deferDeadlineUniversal) { $deferParams.Add('DeferDeadline', $deferDeadlineUniversal) }
                                $promptResult = & $Script:CommandTable."Show-ADTWelcomePrompt$($adtConfig.UI.DialogStyle)" @promptParams @deferParams
                            }
                        }
                        elseif ($welcomeState.RunningProcessDescriptions -or !!$forceCountdown)
                        {
                            # If there is no deferral and processes are running, prompt the user to close running processes with no deferral option.
                            $promptResult = & $Script:CommandTable."Show-ADTWelcomePrompt$($adtConfig.UI.DialogStyle)" @promptParams
                        }
                        else
                        {
                            # If there is no deferral and no processes running, break the while loop.
                            break
                        }

                        # Process the form results.
                        if ($promptResult -eq 'Continue')
                        {
                            # If the user has clicked OK, wait a few seconds for the process to terminate before evaluating the running processes again.
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'The user selected to continue...'
                            if (!$welcomeState.RunningProcesses)
                            {
                                # Break the while loop if there are no processes to close and the user has clicked OK to continue.
                                break
                            }
                            [System.Threading.Thread]::Sleep(2000)
                        }
                        elseif ($promptResult -eq 'Close')
                        {
                            # Force the applications to close.
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'The user selected to force the application(s) to close...'
                            if ($PromptToSave -and $adtEnv.SessionZero -and !$adtEnv.IsProcessUserInteractive)
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Specified [-PromptToSave] option will not be available, because current process is running in session zero and is not interactive.' -Severity 2
                            }

                            # Update the process list right before closing, in case it changed.
                            $PromptToSaveTimeout = [System.TimeSpan]::FromSeconds($adtConfig.UI.PromptToSaveTimeout)
                            foreach ($runningProcess in ($welcomeState.RunningProcesses = & $Script:CommandTable.'Get-ADTRunningProcesses' -ProcessObject $CloseProcesses -InformationAction SilentlyContinue))
                            {
                                # If the PromptToSave parameter was specified and the process has a window open, then prompt the user to save work if there is work to be saved when closing window.
                                if ($PromptToSave -and !($adtEnv.SessionZero -and !$adtEnv.IsProcessUserInteractive) -and ($AllOpenWindowsForRunningProcess = & $Script:CommandTable.'Get-ADTWindowTitle' -ParentProcess $runningProcess.ProcessName -InformationAction SilentlyContinue | & $Script:CommandTable.'Select-Object' -First 1) -and ($runningProcess.MainWindowHandle -ne [IntPtr]::Zero))
                                {
                                    foreach ($OpenWindow in $AllOpenWindowsForRunningProcess)
                                    {
                                        try
                                        {
                                            # Try to bring the window to the front before closing. This doesn't always work.
                                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Stopping process [$($runningProcess.ProcessName)] with window title [$($OpenWindow.WindowTitle)] and prompt to save if there is work to be saved (timeout in [$($adtConfig.UI.PromptToSaveTimeout)] seconds)..."
                                            $null = try
                                            {
                                                [PSADT.GUI.UiAutomation]::BringWindowToFront($OpenWindow.WindowHandle)
                                            }
                                            catch
                                            {
                                                $null
                                            }

                                            # Close out the main window and spin until completion.
                                            if ($runningProcess.CloseMainWindow())
                                            {
                                                $promptToSaveStart = [System.DateTime]::Now
                                                do
                                                {
                                                    if (!($IsWindowOpen = & $Script:CommandTable.'Get-ADTWindowTitle' -WindowHandle $OpenWindow.WindowHandle -InformationAction SilentlyContinue | & $Script:CommandTable.'Select-Object' -First 1))
                                                    {
                                                        break
                                                    }
                                                    [System.Threading.Thread]::Sleep(3000)
                                                }
                                                while (($IsWindowOpen) -and (([System.DateTime]::Now - $promptToSaveStart) -lt $PromptToSaveTimeout))

                                                if ($IsWindowOpen)
                                                {
                                                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Exceeded the [$($adtConfig.UI.PromptToSaveTimeout)] seconds timeout value for the user to save work associated with process [$($runningProcess.ProcessName)] with window title [$($OpenWindow.WindowTitle)]." -Severity 2
                                                }
                                                else
                                                {
                                                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Window [$($OpenWindow.WindowTitle)] for process [$($runningProcess.ProcessName)] was successfully closed."
                                                }
                                            }
                                            else
                                            {
                                                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to call the CloseMainWindow() method on process [$($runningProcess.ProcessName)] with window title [$($OpenWindow.WindowTitle)] because the main window may be disabled due to a modal dialog being shown." -Severity 3
                                            }
                                        }
                                        catch
                                        {
                                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to close window [$($OpenWindow.WindowTitle)] for process [$($runningProcess.ProcessName)].`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
                                        }
                                        finally
                                        {
                                            $runningProcess.Refresh()
                                        }
                                    }
                                }
                                else
                                {
                                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Stopping process $($runningProcess.ProcessName)..."
                                    & $Script:CommandTable.'Stop-Process' -Name $runningProcess.ProcessName -Force -ErrorAction Ignore
                                }
                            }

                            if ($welcomeState.RunningProcesses = & $Script:CommandTable.'Get-ADTRunningProcesses' -ProcessObjects $CloseProcesses -InformationAction SilentlyContinue)
                            {
                                # Apps are still running, give them 2s to close. If they are still running, the Welcome Window will be displayed again.
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Sleeping for 2 seconds because the processes are still not closed...'
                                [System.Threading.Thread]::Sleep(2000)
                            }
                        }
                        elseif ($promptResult -eq 'Timeout')
                        {
                            # Stop the script (if not actioned before the timeout value).
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Installation not actioned before the timeout value.'
                            $BlockExecution = $false
                            if ($adtSession -and (($DeferTimes -ge 0) -or $deferDeadlineUniversal))
                            {
                                & $Script:CommandTable.'Set-ADTDeferHistory' -DeferTimesRemaining $DeferTimes -DeferDeadline $deferDeadlineUniversal
                            }

                            # Dispose the welcome prompt timer here because if we dispose it within the Show-ADTWelcomePrompt function we risk resetting the timer and missing the specified timeout period.
                            if ($welcomeState.WelcomeTimer)
                            {
                                $welcomeState.WelcomeTimer.Dispose()
                                $welcomeState.WelcomeTimer = $null
                            }

                            # Restore minimized windows.
                            if (!$NoMinimizeWindows)
                            {
                                $null = $adtEnv.ShellApp.UndoMinimizeAll()
                            }
                            if ($adtSession)
                            {
                                & $Script:CommandTable.'Close-ADTSession' -ExitCode $adtConfig.UI.DefaultExitCode
                            }
                        }
                        elseif ($promptResult -eq 'Defer')
                        {
                            #  Stop the script (user chose to defer)
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Installation deferred by the user.'
                            $BlockExecution = $false
                            & $Script:CommandTable.'Set-ADTDeferHistory' -DeferTimesRemaining $DeferTimes -DeferDeadline $deferDeadlineUniversal

                            # Restore minimized windows.
                            if (!$NoMinimizeWindows)
                            {
                                $null = $adtEnv.ShellApp.UndoMinimizeAll()
                            }
                            if ($adtSession)
                            {
                                & $Script:CommandTable.'Close-ADTSession' -ExitCode $adtConfig.UI.DeferExitCode
                            }
                        }
                    }
                }

                # Force the processes to close silently, without prompting the user.
                if (($Silent -or ($adtSession -and $adtSession.IsSilent())) -and ($runningProcesses = & $Script:CommandTable.'Get-ADTRunningProcesses' -ProcessObjects $CloseProcesses -InformationAction SilentlyContinue))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Force closing application(s) [$(($runningProcesses.ProcessDescription | & $Script:CommandTable.'Sort-Object' -Unique) -join ',')] without prompting user."
                    $runningProcesses | & $Script:CommandTable.'Stop-Process' -Force -ErrorAction Ignore
                    [System.Threading.Thread]::Sleep(2000)
                }

                # If block execution switch is true, call the function to block execution of these processes.
                if ($BlockExecution -and $CloseProcesses)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message '[-BlockExecution] parameter specified.'
                    & $Script:CommandTable.'Block-ADTAppExecution' -ProcessName $CloseProcesses.Name
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Start-ADTMsiProcess
#
#-----------------------------------------------------------------------------

function Start-ADTMsiProcess
{
    <#
    .SYNOPSIS
        Executes msiexec.exe to perform actions such as install, uninstall, patch, repair, or active setup for MSI and MSP files or MSI product codes.

    .DESCRIPTION
        This function utilizes msiexec.exe to handle various operations on MSI and MSP files, as well as MSI product codes. The operations include installation, uninstallation, patching, repair, and setting up active configurations.

        If the -Action parameter is set to "Install" and the MSI is already installed, the function will terminate without performing any actions.

        The function automatically sets default switches for msiexec based on preferences defined in the config.psd1 file. Additionally, it generates a log file name and creates a verbose log for all msiexec operations, ensuring detailed tracking.

        The MSI or MSP file is expected to reside in the "Files" subdirectory of the App Deploy Toolkit, with transform files expected to be in the same directory as the MSI file.

    .PARAMETER Action
        Specifies the action to be performed. Available options: Install, Uninstall, Patch, Repair, ActiveSetup.

    .PARAMETER FilePath
        The file path to the MSI/MSP file.

    .PARAMETER ProductCode
        The product code of the installed MSI.

    .PARAMETER InstalledApplication
        The InstalledApplication object of the installed MSI.

    .PARAMETER Transforms
        The name(s) of the transform file(s) to be applied to the MSI. The transform files should be in the same directory as the MSI file.

    .PARAMETER Patches
        The name(s) of the patch (MSP) file(s) to be applied to the MSI for the "Install" action. The patch files should be in the same directory as the MSI file.

    .PARAMETER ArgumentList
        Overrides the default parameters specified in the config.psd1 file. The install default is: "REBOOT=ReallySuppress /QB!". The uninstall default is: "REBOOT=ReallySuppress /QN".

    .PARAMETER AdditionalArgumentList
        Adds additional parameters to the default set specified in the config.psd1 file. The install default is: "REBOOT=ReallySuppress /QB!". The uninstall default is: "REBOOT=ReallySuppress /QN".

    .PARAMETER SecureArgumentList
        Hides all parameters passed to the MSI or MSP file from the toolkit log file.

    .PARAMETER LoggingOptions
        Overrides the default logging options specified in the config.psd1 file.

    .PARAMETER LogFileName
        Overrides the default log file name. The default log file name is generated from the MSI file name. If LogFileName does not end in .log, it will be automatically appended.

        For uninstallations, by default the product code is resolved to the DisplayName and version of the application.

    .PARAMETER WorkingDirectory
        Overrides the working directory. The working directory is set to the location of the MSI file.

    .PARAMETER SkipMSIAlreadyInstalledCheck
        Skips the check to determine if the MSI is already installed on the system. Default is: $false.

    .PARAMETER IncludeUpdatesAndHotfixes
        Include matches against updates and hotfixes in results.

    .PARAMETER NoWait
        Immediately continue after executing the process.

    .PARAMETER PassThru
        Returns ExitCode, STDOut, and STDErr output from the process.

    .PARAMETER SuccessExitCodes
        List of exit codes to be considered successful. Defaults to values set during ADTSession initialization, otherwise: 0

    .PARAMETER RebootExitCodes
        List of exit codes to indicate a reboot is required. Defaults to values set during ADTSession initialization, otherwise: 1641, 3010

    .PARAMETER IgnoreExitCodes
        List the exit codes to ignore or * to ignore all exit codes.

    .PARAMETER PriorityClass
        Specifies priority class for the process. Options: Idle, Normal, High, AboveNormal, BelowNormal, RealTime. Default: Normal

    .PARAMETER RepairFromSource
        Specifies whether we should repair from source. Also rewrites local cache. Default: $false

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.ProcessResult

        Returns an object with the results of the installation if -PassThru is specified.
        - ExitCode
        - StdOut
        - StdErr

    .EXAMPLE
        Start-ADTMsiProcess -Action 'Install' -FilePath 'Adobe_FlashPlayer_11.2.202.233_x64_EN.msi'

        Install an MSI.

    .EXAMPLE
        Start-ADTMsiProcess -Action 'Install' -FilePath 'Adobe_FlashPlayer_11.2.202.233_x64_EN.msi' -Transforms 'Adobe_FlashPlayer_11.2.202.233_x64_EN_01.mst' -ArgumentList '/QN'

        Install an MSI, applying a transform and overriding the default MSI toolkit parameters.

    .EXAMPLE
        $ExecuteMSIResult = Start-ADTMsiProcess -Action 'Install' -FilePath 'Adobe_FlashPlayer_11.2.202.233_x64_EN.msi' -PassThru

        Install an MSI and stores the result of the execution into a variable by using the -PassThru option.

    .EXAMPLE
        Start-ADTMsiProcess -Action 'Uninstall' -ProductCode '{26923b43-4d38-484f-9b9e-de460746276c}'

        Uninstall an MSI using a product code.

    .EXAMPLE
        Start-ADTMsiProcess -Action 'Patch' -FilePath 'Adobe_Reader_11.0.3_EN.msp'

        Install an MSP.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Int32])]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateSet('Install', 'Uninstall', 'Patch', 'Repair', 'ActiveSetup')]
        [System.String]$Action = 'Install',

        [Parameter(Mandatory = $true, ParameterSetName = 'FilePath', ValueFromPipeline = $true, HelpMessage = 'Please supply the path to the MSI/MSP file to process.')]
        [ValidateScript({
                if ([System.IO.Path]::GetExtension($_) -notmatch '^\.ms[ip]$')
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified input has an invalid file extension.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath,

        [Parameter(Mandatory = $true, ParameterSetName = 'ProductCode', ValueFromPipeline = $true, HelpMessage = 'Please supply the Product Code to process.')]
        [ValidateNotNullOrEmpty()]
        [System.Guid]$ProductCode,

        [Parameter(Mandatory = $true, ParameterSetName = 'InstalledApplication', ValueFromPipeline = $true, HelpMessage = 'Please supply the InstalledApplication object to process.')]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.InstalledApplication]$InstalledApplication,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Transforms,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$AdditionalArgumentList,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Patches,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LoggingOptions,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogFileName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory,

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
        [System.Int32[]]$SuccessExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]$RebootExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$IgnoreExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Diagnostics.ProcessPriorityClass]$PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$RepairFromSource
    )

    begin
    {
        # The use of a ProductCode with an Install action is not supported.
        if ($ProductCode -and ($Action -eq 'Install'))
        {
            $naerParams = @{
                Exception = [System.InvalidOperationException]::new("The ProductCode parameter can only be used with non-install actions.")
                Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                ErrorId = 'ProductCodeInstallActionNotSupported'
                TargetObject = $PSBoundParameters
                RecommendedAction = "Please review the supplied parameters and try again."
            }
            $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
        }
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet; $adtConfig = & $Script:CommandTable.'Get-ADTConfig'
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Determine whether the input is a ProductCode or not.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Executing MSI action [$Action]..."

                # If the MSI is in the Files directory, set the full path to the MSI.
                $msiProduct = switch ($PSCmdlet.ParameterSetName)
                {
                    FilePath
                    {
                        if (& $Script:CommandTable.'Test-Path' -LiteralPath $FilePath -PathType Leaf)
                        {
                            (& $Script:CommandTable.'Get-Item' -LiteralPath $FilePath).FullName
                        }
                        elseif ($adtSession -and [System.IO.File]::Exists(($dirFilesPath = [System.IO.Path]::Combine($adtSession.DirFiles, $FilePath))))
                        {
                            $dirFilesPath
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to find the file [$FilePath]." -Severity 3
                            $naerParams = @{
                                Exception = [System.IO.FileNotFoundException]::new("Failed to find the file [$FilePath].")
                                Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                                ErrorId = 'FilePathNotFound'
                                TargetObject = $FilePath
                                RecommendedAction = "Please confirm the path of the file and try again."
                            }
                            throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                        }
                        break
                    }

                    ProductCode
                    {
                        $ProductCode.ToString('B')
                        break
                    }

                    InstalledApplication
                    {
                        $InstalledApplication.ProductCode.ToString('B')
                        break
                    }
                }

                # Fix up any bad file paths.
                if ([System.IO.Path]::GetExtension($msiProduct) -eq '.msi')
                {
                    # Iterate transforms.
                    if ($Transforms)
                    {
                        for ($i = 0; $i -lt $Transforms.Length; $i++)
                        {
                            if ([System.IO.File]::Exists(($fullPath = & $Script:CommandTable.'Join-Path' -Path (& $Script:CommandTable.'Split-Path' -Path $msiProduct -Parent) -ChildPath $Transforms[$i].Replace('.\', ''))))
                            {
                                $Transforms[$i] = $fullPath
                            }
                        }
                    }

                    # Iterate patches.
                    if ($Patches)
                    {
                        for ($i = 0; $i -lt $Patches.Length; $i++)
                        {
                            if ([System.IO.File]::Exists(($fullPath = & $Script:CommandTable.'Join-Path' -Path (& $Script:CommandTable.'Split-Path' -Path $msiProduct -Parent) -ChildPath $Patches[$i].Replace('.\', ''))))
                            {
                                $Patches[$i] = $fullPath
                            }
                        }
                    }
                }

                # If the provided MSI was a file path, get the Property table and store it.
                $msiPropertyTable = if ([System.IO.Path]::GetExtension($msiProduct) -eq '.msi')
                {
                    $gmtpParams = @{ Path = $msiProduct; Table = 'Property' }; if ($Transforms) { $gmtpParams.Add('TransformPath', $transforms) }
                    & $Script:CommandTable.'Get-ADTMsiTableProperty' @gmtpParams
                }

                # Get the ProductCode of the MSI.
                $msiProductCode = if ($ProductCode)
                {
                    $ProductCode
                }
                elseif ($InstalledApplication)
                {
                    $InstalledApplication.ProductCode
                }
                elseif ($msiPropertyTable)
                {
                    $msiPropertyTable.ProductCode
                }

                # Check if the MSI is already installed. If no valid ProductCode to check or SkipMSIAlreadyInstalledCheck supplied, then continue with requested MSI action.
                $msiInstalled = if ($msiProductCode -and !$SkipMSIAlreadyInstalledCheck)
                {
                    if (!$InstalledApplication -and ($installedApps = & $Script:CommandTable.'Get-ADTApplication' -FilterScript { $_.ProductCode -eq $msiProductCode } -IncludeUpdatesAndHotfixes:$IncludeUpdatesAndHotfixes))
                    {
                        $InstalledApplication = $installedApps
                    }
                    !!$InstalledApplication
                }
                else
                {
                    $Action -ne 'Install'
                }

                # Return early if we're installing an installed product, or anything else for a non-installed product.
                if ($msiInstalled -and ($Action -eq 'Install'))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The MSI is already installed on this system. Skipping action [$Action]..."
                    return $(if ($PassThru) { [PSADT.Types.ProcessResult]::new(1638, $null, $null) })
                }
                elseif (!$msiInstalled -and ($Action -ne 'Install'))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The MSI is not installed on this system. Skipping action [$Action]..."
                    return
                }

                # Set up the log file to use.
                $logFile = if ($PSBoundParameters.ContainsKey('LogFileName'))
                {
                    [System.IO.Path]::GetFileNameWithoutExtension($LogFileName)
                }
                elseif ($InstalledApplication)
                {
                    (& $Script:CommandTable.'Remove-ADTInvalidFileNameChars' -Name ($InstalledApplication.DisplayName + '_' + $InstalledApplication.DisplayVersion)) -replace '\s+'
                }
                elseif ($msiPropertyTable)
                {
                    (& $Script:CommandTable.'Remove-ADTInvalidFileNameChars' -Name ($msiPropertyTable.ProductName + '_' + $msiPropertyTable.ProductVersion)) -replace '\s+'
                }

                # Build the log path to use.
                $logPath = if ($logFile)
                {
                    if ($adtSession -and $adtConfig.Toolkit.CompressLogs)
                    {
                        & $Script:CommandTable.'Join-Path' -Path $adtSession.LogTempFolder -ChildPath $logFile
                    }
                    else
                    {
                        # Create the Log directory if it doesn't already exist.
                        if (![System.IO.Directory]::Exists($adtConfig.MSI.LogPath))
                        {
                            $null = [System.IO.Directory]::CreateDirectory($adtConfig.MSI.LogPath)
                        }

                        # Build the log file path.
                        & $Script:CommandTable.'Join-Path' -Path $adtConfig.MSI.LogPath -ChildPath $logFile
                    }
                }

                # Set the installation parameters.
                if ($adtSession -and $adtSession.IsNonInteractive())
                {
                    $msiInstallDefaultParams = $adtConfig.MSI.SilentParams
                    $msiUninstallDefaultParams = $adtConfig.MSI.SilentParams
                }
                else
                {
                    $msiInstallDefaultParams = $adtConfig.MSI.InstallParams
                    $msiUninstallDefaultParams = $adtConfig.MSI.UninstallParams
                }

                # Build the MSI parameters.
                switch ($action)
                {
                    Install
                    {
                        $option = '/i'
                        $msiLogFile = if ($logPath) { "$($logPath)_$($_)" }
                        $msiDefaultParams = $msiInstallDefaultParams
                        break
                    }
                    Uninstall
                    {
                        $option = '/x'
                        $msiLogFile = if ($logPath) { "$($logPath)_$($_)" }
                        $msiDefaultParams = $msiUninstallDefaultParams
                        break
                    }
                    Patch
                    {
                        $option = '/update'
                        $msiLogFile = if ($logPath) { "$($logPath)_$($_)" }
                        $msiDefaultParams = $msiInstallDefaultParams
                        break
                    }
                    Repair
                    {
                        $option = "/f$(if ($RepairFromSource) {'vomus'})"
                        $msiLogFile = if ($logPath) { "$($logPath)_$($_)" }
                        $msiDefaultParams = $msiInstallDefaultParams
                        break
                    }
                    ActiveSetup
                    {
                        $option = '/fups'
                        $msiLogFile = if ($logPath) { "$($logPath)_$($_)" }
                        $msiDefaultParams = $null
                        break
                    }
                }

                # Post-process the MSI log file variable.
                if ($msiLogFile)
                {
                    # Append the username to the log file name if the toolkit is not running as an administrator, since users do not have the rights to modify files in the ProgramData folder that belong to other users.
                    if (!(& $Script:CommandTable.'Test-ADTCallerIsAdmin'))
                    {
                        $msiLogFile = $msiLogFile + '_' + (& $Script:CommandTable.'Remove-ADTInvalidFileNameChars' -Name ([System.Environment]::UserName))
                    }

                    # Append ".log" to the MSI logfile path and enclose in quotes.
                    if ([IO.Path]::GetExtension($msiLogFile) -ne '.log')
                    {
                        $msiLogFile = "`"$($msiLogFile + '.log')`""
                    }
                }

                # Set the working directory of the MSI.
                if ($PSCmdlet.ParameterSetName.Equals('FilePath') -and !$workingDirectory)
                {
                    $WorkingDirectory = [System.IO.Path]::GetDirectoryName($msiProduct)
                }

                # Enumerate all transforms specified, qualify the full path if possible and enclose in quotes.
                $mstFile = if ($Transforms)
                {
                    "`"$($Transforms -join ';')`""
                }

                # Enumerate all patches specified, qualify the full path if possible and enclose in quotes.
                $mspFile = if ($Patches)
                {
                    "`"$($Patches -join ';')`""
                }

                # Start building the MsiExec command line starting with the base action and file.
                $argsMSI = "$option `"$msiProduct`""

                # Add MST.
                if ($mstFile)
                {
                    $argsMSI = "$argsMSI TRANSFORMS=$mstFile TRANSFORMSSECURE=1"
                }

                # Add MSP.
                if ($mspFile)
                {
                    $argsMSI = "$argsMSI PATCH=$mspFile"
                }

                # Replace default parameters if specified.
                $argsMSI = if ($ArgumentList)
                {
                    "$argsMSI $([System.String]::Join(' ', $ArgumentList))"
                }
                else
                {
                    "$argsMSI $msiDefaultParams"
                }

                # Add reinstallmode and reinstall variable for Patch.
                if ($action -eq 'Patch')
                {
                    $argsMSI = "$argsMSI REINSTALLMODE=ecmus REINSTALL=ALL"
                }

                # Append parameters to default parameters if specified.
                if ($AdditionalArgumentList)
                {
                    $argsMSI = "$argsMSI $([System.String]::Join(' ', $AdditionalArgumentList))"
                }

                # Add custom Logging Options if specified, otherwise, add default Logging Options from Config file.
                if ($msiLogFile)
                {
                    $argsMSI = if ($LoggingOptions)
                    {
                        "$argsMSI $LoggingOptions $msiLogFile"
                    }
                    else
                    {
                        "$argsMSI $($adtConfig.MSI.LoggingOptions) $msiLogFile"
                    }
                }

                # Build the hashtable with the options that will be passed to Start-ADTProcess using splatting.
                $ExecuteProcessSplat = @{
                    FilePath = "$([System.Environment]::SystemDirectory)\msiexec.exe"
                    ArgumentList = $argsMSI
                    WindowStyle = 'Normal'
                }
                if ($WorkingDirectory)
                {
                    $ExecuteProcessSplat.Add('WorkingDirectory', $WorkingDirectory)
                }
                if ($SecureArgumentList)
                {
                    $ExecuteProcessSplat.Add('SecureArgumentList', $SecureArgumentList)
                }
                if ($PassThru)
                {
                    $ExecuteProcessSplat.Add('PassThru', $PassThru)
                }
                if ($SuccessExitCodes)
                {
                    $ExecuteProcessSplat.Add('SuccessExitCodes', $SuccessExitCodes)
                }
                if ($RebootExitCodes)
                {
                    $ExecuteProcessSplat.Add('RebootExitCodes', $RebootExitCodes)
                }
                if ($IgnoreExitCodes)
                {
                    $ExecuteProcessSplat.Add('IgnoreExitCodes', $IgnoreExitCodes)
                }
                if ($PriorityClass)
                {
                    $ExecuteProcessSplat.Add('PriorityClass', $PriorityClass)
                }
                if ($NoWait)
                {
                    $ExecuteProcessSplat.Add('NoWait', $NoWait)
                }

                # Call the Start-ADTProcess function.
                $result = & $Script:CommandTable.'Start-ADTProcess' @ExecuteProcessSplat

                # Refresh environment variables for Windows Explorer process as Windows does not consistently update environment variables created by MSIs.
                & $Script:CommandTable.'Update-ADTDesktop'

                # Return the results if passing through.
                if ($PassThru -and $result)
                {
                    return $result
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Start-ADTMspProcess
#
#-----------------------------------------------------------------------------

function Start-ADTMspProcess
{
    <#
    .SYNOPSIS
        Executes an MSP file using the same logic as Start-ADTMsiProcess.

    .DESCRIPTION
        Reads SummaryInfo targeted product codes in MSP file and determines if the MSP file applies to any installed products. If a valid installed product is found, triggers the Start-ADTMsiProcess function to patch the installation.

        Uses default config MSI parameters. You can use -AdditionalArgumentList to add additional parameters.

    .PARAMETER FilePath
        Path to the MSP file.

    .PARAMETER AdditionalArgumentList
        Additional parameters.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Start-ADTMspProcess -FilePath 'Adobe_Reader_11.0.3_EN.msp'

        Executes the specified MSP file for Adobe Reader 11.0.3.

    .EXAMPLE
        Start-ADTMspProcess -FilePath 'AcroRdr2017Upd1701130143_MUI.msp' -AdditionalArgumentList 'ALLUSERS=1'

        Executes the specified MSP file for Acrobat Reader 2017 with additional parameters.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Int32])]
    param
    (
        [Parameter(Mandatory = $true, HelpMessage = 'Please supply the path to the MSP file to process.')]
        [ValidateScript({
                if ([System.IO.Path]::GetExtension($_) -notmatch '^\.msp$')
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified input has an invalid file extension.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$AdditionalArgumentList
    )

    begin
    {
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # If the MSP is in the Files directory, set the full path to the MSP.
                $mspFile = if ($adtSession -and [System.IO.File]::Exists(($dirFilesPath = [System.IO.Path]::Combine($adtSession.DirFiles, $FilePath))))
                {
                    $dirFilesPath
                }
                elseif (& $Script:CommandTable.'Test-Path' -LiteralPath $FilePath)
                {
                    (& $Script:CommandTable.'Get-Item' -LiteralPath $FilePath).FullName
                }
                else
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Failed to find MSP file [$FilePath]." -Severity 3
                    $naerParams = @{
                        Exception = [System.IO.FileNotFoundException]::new("Failed to find MSP file [$FilePath].")
                        Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                        ErrorId = 'MsiFileNotFound'
                        TargetObject = $FilePath
                        RecommendedAction = "Please confirm the path of the MSP file and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }

                # Create a Windows Installer object and open the database in read-only mode.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Checking MSP file for valid product codes.'
                [__ComObject]$Installer = & $Script:CommandTable.'New-Object' -ComObject WindowsInstaller.Installer
                [__ComObject]$Database = & $Script:CommandTable.'Invoke-ADTObjectMethod' -InputObject $Installer -MethodName OpenDatabase -ArgumentList @($mspFile, 32)

                # Get the SummaryInformation from the Windows Installer database and store all product codes found.
                [__ComObject]$SummaryInformation = & $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $Database -PropertyName SummaryInformation
                $AllTargetedProductCodes = & $Script:CommandTable.'Get-ADTApplication' -ProductCode (& $Script:CommandTable.'Get-ADTObjectProperty' -InputObject $SummaryInformation -PropertyName Property -ArgumentList @(7)).Split(';')

                # Free our COM objects.
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($SummaryInformation)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Database)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Installer)

                # If the application is installed, patch it.
                if ($AllTargetedProductCodes)
                {
                    & $Script:CommandTable.'Start-ADTMsiProcess' -Action Patch @PSBoundParameters
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Start-ADTProcess
#
#-----------------------------------------------------------------------------

function Start-ADTProcess
{
    <#
    .SYNOPSIS
        Execute a process with optional arguments, working directory, window style.

    .DESCRIPTION
        Executes a process, e.g. a file included in the Files directory of the App Deploy Toolkit, or a file on the local machine. Provides various options for handling the return codes (see Parameters).

    .PARAMETER FilePath
        Path to the file to be executed. If the file is located directly in the "Files" directory of the App Deploy Toolkit, only the file name needs to be specified.

        Otherwise, the full path of the file must be specified. If the files is in a subdirectory of "Files", use the "$($adtSession.DirFiles)" variable as shown in the example.

    .PARAMETER ArgumentList
        Arguments to be passed to the executable.

    .PARAMETER SecureArgumentList
        Hides all parameters passed to the executable from the Toolkit log file.

    .PARAMETER WindowStyle
        Style of the window of the process executed. Options: Normal, Hidden, Maximized, Minimized. Default: Normal. Only works for native Windows GUI applications. If the WindowStyle is set to Hidden, UseShellExecute should be set to $true.

        Note: Not all processes honor WindowStyle. WindowStyle is a recommendation passed to the process. They can choose to ignore it.

    .PARAMETER CreateNoWindow
        Specifies whether the process should be started with a new window to contain it. Only works for Console mode applications. UseShellExecute should be set to $false. Default is false.

    .PARAMETER WorkingDirectory
        The working directory used for executing the process. Defaults to the directory of the file being executed. The use of UseShellExecute affects this parameter.

    .PARAMETER NoWait
        Immediately continue after executing the process.

    .PARAMETER PassThru
        If NoWait is not specified, returns an object with ExitCode, STDOut and STDErr output from the process. If NoWait is specified, returns an object with Id, Handle and ProcessName.

    .PARAMETER WaitForMsiExec
        Sometimes an EXE bootstrapper will launch an MSI install. In such cases, this variable will ensure that this function waits for the msiexec engine to become available before starting the install.

    .PARAMETER MsiExecWaitTime
        Specify the length of time in seconds to wait for the msiexec engine to become available. Default: 600 seconds (10 minutes).

    .PARAMETER SuccessExitCodes
        List of exit codes to be considered successful. Defaults to values set during ADTSession initialization, otherwise: 0

    .PARAMETER RebootExitCodes
        List of exit codes to indicate a reboot is required. Defaults to values set during ADTSession initialization, otherwise: 1641, 3010

    .PARAMETER IgnoreExitCodes
        List the exit codes to ignore or * to ignore all exit codes.

    .PARAMETER PriorityClass
        Specifies priority class for the process. Options: Idle, Normal, High, AboveNormal, BelowNormal, RealTime. Default: Normal

    .PARAMETER UseShellExecute
        Specifies whether to use the operating system shell to start the process. $true if the shell should be used when starting the process; $false if the process should be created directly from the executable file.

        The word "Shell" in this context refers to a graphical shell (similar to the Windows shell) rather than command shells (for example, bash or sh) and lets users launch graphical applications or open documents. It lets you open a file or a url and the Shell will figure out the program to open it with.

        The WorkingDirectory property behaves differently depending on the value of the UseShellExecute property. When UseShellExecute is true, the WorkingDirectory property specifies the location of the executable. When UseShellExecute is false, the WorkingDirectory property is not used to find the executable. Instead, it is used only by the process that is started and has meaning only within the context of the new process.

        If you set UseShellExecute to $true, there will be no available output from the process.

    .EXAMPLE
        Start-ADTProcess -FilePath 'setup.exe' -ArgumentList '/S' -IgnoreExitCodes 1,2

        Launch InstallShield "setup.exe" from the ".\Files" sub-directory.

    .EXAMPLE
        Start-ADTProcess -FilePath "$($adtSession.DirFiles)\Bin\setup.exe" -ArgumentList '/S' -WindowStyle 'Hidden'

        Launch InstallShield "setup.exe" from the ".\Files\Bin" sub-directory.

    .EXAMPLE
        Start-ADTProcess -FilePath 'uninstall_flash_player_64bit.exe' -ArgumentList '/uninstall' -WindowStyle 'Hidden'

        If the file is in the "Files" directory of the AppDeployToolkit, only the file name needs to be specified.

    .EXAMPLE
        Start-ADTProcess -FilePath 'setup.exe' -ArgumentList "-s -f2`"$((Get-ADTConfig).Toolkit.LogPath)\$($adtSession.InstallName).log`""

        Launch InstallShield "setup.exe" from the ".\Files" sub-directory and force log files to the logging folder.

    .EXAMPLE
        Start-ADTProcess -FilePath 'setup.exe' -ArgumentList "/s /v`"ALLUSERS=1 /qn /L* `"$((Get-ADTConfig).Toolkit.LogPath)\$($adtSession.InstallName).log`"`""

        Launch InstallShield "setup.exe" with embedded MSI and force log files to the logging folder.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.ProcessResult

        Returns an object with the results of the installation if -PassThru is specified.
        - ExitCode
        - StdOut
        - StdErr

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([PSADT.Types.ProcessResult])]
    [OutputType([PSADT.Types.ProcessInfo])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ArgumentList,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Diagnostics.ProcessWindowStyle]$WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$CreateNoWindow,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$NoWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$WaitForMsiExec,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$MsiExecWaitTime,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]$SuccessExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]$RebootExitCodes,

        [Parameter(Mandatory = $false)]
        [SupportsWildcards()]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$IgnoreExitCodes,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Diagnostics.ProcessPriorityClass]$PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$UseShellExecute
    )

    begin
    {
        # Initalize function and get required objects.
        $adtSession = & $Script:CommandTable.'Initialize-ADTModuleIfUnitialized' -Cmdlet $PSCmdlet
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Set up defaults if not specified.
        if (!$PSBoundParameters.ContainsKey('MsiExecWaitTime'))
        {
            $MsiExecWaitTime = (& $Script:CommandTable.'Get-ADTConfig').MSI.MutexWaitTime
        }
        if (!$PSBoundParameters.ContainsKey('SuccessExitCodes'))
        {
            $SuccessExitCodes = if ($adtSession)
            {
                $adtSession.AppSuccessExitCodes
            }
            else
            {
                0
            }
        }
        if (!$PSBoundParameters.ContainsKey('RebootExitCodes'))
        {
            $RebootExitCodes = if ($adtSession)
            {
                $adtSession.AppRebootExitCodes
            }
            else
            {
                1641, 3010
            }
        }

        # Set up initial variables.
        $extInvoker = !(& $Script:CommandTable.'Get-PSCallStack')[1].InvocationInfo.MyCommand.Source.StartsWith($MyInvocation.MyCommand.Module.Name)
        $stdOutBuilder = [System.Text.StringBuilder]::new()
        $stdErrBuilder = [System.Text.StringBuilder]::new()
        $stdOutEvent = $stdErrEvent = $null
        $stdOut = $stdErr = $null
        $returnCode = $null
    }

    process
    {
        try
        {
            try
            {
                # Validate and find the fully qualified path for the $FilePath variable.
                if ([System.IO.Path]::IsPathRooted($FilePath) -and [System.IO.Path]::HasExtension($FilePath))
                {
                    if (![System.IO.File]::Exists($FilePath))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "File [$FilePath] not found." -Severity 3
                        $naerParams = @{
                            Exception = [System.IO.FileNotFoundException]::new("File [$FilePath] not found.")
                            Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                            ErrorId = 'PathFileNotFound'
                            TargetObject = $FilePath
                            RecommendedAction = "Please confirm the path of the specified file and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "[$FilePath] is a valid fully qualified path, continue."
                }
                else
                {
                    # Get the fully qualified path for the file using DirFiles, the current directory, then the system's path environment variable.
                    if (!($fqPath = & $Script:CommandTable.'Get-Item' -Path ("$(if ($adtSession) { "$($adtSession.DirFiles);" })$($ExecutionContext.SessionState.Path.CurrentLocation.Path);$([System.Environment]::GetEnvironmentVariable('PATH'))".TrimEnd(';').Split(';').TrimEnd('\') -replace '$', "\$FilePath") -ErrorAction Ignore | & $Script:CommandTable.'Select-Object' -ExpandProperty FullName -First 1))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "[$FilePath] contains an invalid path or file name." -Severity 3
                        $naerParams = @{
                            Exception = [System.IO.FileNotFoundException]::new("[$FilePath] contains an invalid path or file name.")
                            Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                            ErrorId = 'PathFileNotFound'
                            TargetObject = $FilePath
                            RecommendedAction = "Please confirm the path of the specified file and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "[$FilePath] successfully resolved to fully qualified path [$fqPath]."
                    $FilePath = $fqPath
                }

                # Set the Working directory if not specified.
                if (!$WorkingDirectory)
                {
                    $WorkingDirectory = [System.IO.Path]::GetDirectoryName($FilePath)
                }

                # If the WindowStyle parameter is set to 'Hidden', set the UseShellExecute parameter to '$true' unless specifically specified.
                if ($WindowStyle.Equals([System.Diagnostics.ProcessWindowStyle]::Hidden) -and !$PSBoundParameters.ContainsKey('UseShellExecute'))
                {
                    $UseShellExecute = $true
                }

                # If MSI install, check to see if the MSI installer service is available or if another MSI install is already underway.
                # Please note that a race condition is possible after this check where another process waiting for the MSI installer
                # to become available grabs the MSI Installer mutex before we do. Not too concerned about this possible race condition.
                if (($FilePath -match 'msiexec') -or $WaitForMsiExec)
                {
                    $MsiExecAvailable = & $Script:CommandTable.'Test-ADTMutexAvailability' -MutexName 'Global\_MSIExecute' -MutexWaitTime ([System.TimeSpan]::FromSeconds($MsiExecWaitTime))
                    [System.Threading.Thread]::Sleep(1000)
                    if (!$MsiExecAvailable)
                    {
                        # Default MSI exit code for install already in progress.
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Another MSI installation is already in progress and needs to be completed before proceeding with this installation.' -Severity 3
                        $returnCode = 1618
                        $naerParams = @{
                            Exception = [System.TimeoutException]::new('Another MSI installation is already in progress and needs to be completed before proceeding with this installation.')
                            Category = [System.Management.Automation.ErrorCategory]::ResourceBusy
                            ErrorId = 'MsiExecUnavailable'
                            TargetObject = $FilePath
                            RecommendedAction = "Please wait for the current MSI operation to finish and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }
                }

                try
                {
                    # Disable Zone checking to prevent warnings when running executables.
                    [System.Environment]::SetEnvironmentVariable('SEE_MASK_NOZONECHECKS', 1)

                    # Define process.
                    $process = [System.Diagnostics.Process]@{
                        StartInfo = [System.Diagnostics.ProcessStartInfo]@{
                            FileName = $FilePath
                            WorkingDirectory = $WorkingDirectory
                            UseShellExecute = $UseShellExecute
                            ErrorDialog = $false
                            RedirectStandardOutput = $true
                            RedirectStandardError = $true
                            CreateNoWindow = $CreateNoWindow
                            WindowStyle = $WindowStyle
                        }
                    }
                    if ($ArgumentList)
                    {
                        $process.StartInfo.Arguments = $ArgumentList
                    }
                    if ($process.StartInfo.UseShellExecute)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'UseShellExecute is set to true, standard output and error will not be available.'
                        $process.StartInfo.RedirectStandardOutput = $false
                        $process.StartInfo.RedirectStandardError = $false
                    }
                    else
                    {
                        # Add event handler to capture process's standard output redirection.
                        $processEventHandler = { $Event.MessageData.AppendLine($(if (![System.String]::IsNullOrWhiteSpace($EventArgs.Data)) { $EventArgs.Data })) }
                        $stdOutEvent = & $Script:CommandTable.'Register-ObjectEvent' -InputObject $process -Action $processEventHandler -EventName OutputDataReceived -MessageData $stdOutBuilder
                        $stdErrEvent = & $Script:CommandTable.'Register-ObjectEvent' -InputObject $process -Action $processEventHandler -EventName ErrorDataReceived -MessageData $stdErrBuilder
                    }

                    # Start Process.
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Working Directory is [$WorkingDirectory]."
                    if ($ArgumentList)
                    {
                        if ($SecureArgumentList)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Executing [$FilePath (Parameters Hidden)]..."
                        }
                        elseif ($ArgumentList -match '-Command \&')
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Executing [$FilePath [PowerShell ScriptBlock]]..."
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Executing [$FilePath $ArgumentList]..."
                        }
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Executing [$FilePath]..."
                    }
                    $null = $process.Start()

                    # Set priority
                    if ($PriorityClass -ne 'Normal')
                    {
                        try
                        {
                            if (!$process.HasExited)
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Changing the priority class for the process to [$PriorityClass]"
                                $process.PriorityClass = $PriorityClass
                            }
                            else
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Cannot change the priority class for the process to [$PriorityClass], because the process has exited already." -Severity 2
                            }
                        }
                        catch
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Failed to change the priority class for the process.' -Severity 2
                        }
                    }

                    # NoWait specified, return process details. If it isn't specified, start reading standard Output and Error streams.
                    if ($NoWait)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'NoWait parameter specified. Continuing without waiting for exit code...'
                        if ($PassThru)
                        {
                            if (!$process.HasExited)
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'PassThru parameter specified, returning process details object.'
                                $PSCmdlet.WriteObject([PSADT.Types.ProcessInfo]::new(
                                        $process.Id,
                                        $process.Handle,
                                        $process.ProcessName
                                    ))
                            }
                            else
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'PassThru parameter specified, however the process has already exited.'
                            }
                        }
                    }
                    else
                    {
                        # Read all streams to end and wait for the process to exit.
                        if (!$process.StartInfo.UseShellExecute)
                        {
                            $process.BeginOutputReadLine()
                            $process.BeginErrorReadLine()
                        }
                        $process.WaitForExit()

                        # HasExited indicates that the associated process has terminated, either normally or abnormally. Wait until HasExited returns $true.
                        while (!$process.HasExited)
                        {
                            $process.Refresh()
                            [System.Threading.Thread]::Sleep(1000)
                        }

                        # Get the exit code for the process.
                        $returnCode = $process.ExitCode

                        # Process all streams.
                        if (!$process.StartInfo.UseShellExecute)
                        {
                            # Unregister standard output and error event to retrieve process output.
                            if ($stdOutEvent)
                            {
                                & $Script:CommandTable.'Unregister-Event' -SourceIdentifier $stdOutEvent.Name
                                $stdOutEvent = $null
                            }
                            if ($stdErrEvent)
                            {
                                & $Script:CommandTable.'Unregister-Event' -SourceIdentifier $stdErrEvent.Name
                                $stdErrEvent = $null
                            }
                            $stdOut = $stdOutBuilder.ToString().Trim()
                            $stdErr = $stdErrBuilder.ToString().Trim()
                            if (![System.String]::IsNullOrWhiteSpace($stdErr))
                            {
                                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Standard error output from the process: $stdErr" -Severity 3
                            }
                        }
                    }
                }
                catch
                {
                    throw
                }
                finally
                {
                    # Make sure the standard output and error event is unregistered.
                    if ($process.StartInfo.UseShellExecute -eq $false)
                    {
                        if ($stdOutEvent)
                        {
                            & $Script:CommandTable.'Unregister-Event' -SourceIdentifier $stdOutEvent.Name -ErrorAction Ignore
                            $stdOutEvent = $null
                        }
                        if ($stdErrEvent)
                        {
                            & $Script:CommandTable.'Unregister-Event' -SourceIdentifier $stdErrEvent.Name -ErrorAction Ignore
                            $stdErrEvent = $null
                        }
                    }

                    # Free resources associated with the process, this does not cause process to exit.
                    if ($process)
                    {
                        $process.Dispose()
                    }

                    # Re-enable zone checking.
                    [System.Environment]::SetEnvironmentVariable('SEE_MASK_NOZONECHECKS', $null)
                }

                if (!$NoWait)
                {
                    # Open variable to store the error message if we failed as we need it when we're determining whether we throw or not.
                    $errorMessage = $null

                    # Check to see whether we should ignore exit codes.
                    if ($IgnoreExitCodes -and ($($IgnoreExitCodes).Equals('*') -or ([System.Int32[]]$IgnoreExitCodes).Contains($returnCode)))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Execution completed and the exit code [$returnCode] is being ignored."
                    }
                    elseif ($RebootExitCodes.Contains($returnCode))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Execution completed successfully with exit code [$returnCode]. A reboot is required." -Severity 2
                    }
                    elseif (($returnCode -eq 1605) -and ($FilePath -match 'msiexec'))
                    {
                        $errorMessage = "Execution failed with exit code [$returnCode] because the product is not currently installed."
                    }
                    elseif (($returnCode -eq -2145124329) -and ($FilePath -match 'wusa'))
                    {
                        $errorMessage = "Execution failed with exit code [$returnCode] because the Windows Update is not applicable to this system."
                    }
                    elseif (($returnCode -eq 17025) -and ($FilePath -match 'fullfile'))
                    {
                        $errorMessage = "Execution failed with exit code [$returnCode] because the Office Update is not applicable to this system."
                    }
                    elseif ($SuccessExitCodes.Contains($returnCode))
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Execution completed successfully with exit code [$returnCode]." -Severity 0
                    }
                    else
                    {
                        if (($MsiExitCodeMessage = if ($FilePath -match 'msiexec') { & $Script:CommandTable.'Get-ADTMsiExitCodeMessage' -MsiExitCode $returnCode }))
                        {
                            $errorMessage = "Execution failed with exit code [$returnCode]: $MsiExitCodeMessage"
                        }
                        else
                        {
                            $errorMessage = "Execution failed with exit code [$returnCode]."
                        }
                    }

                    # Generate and store the PassThru data.
                    $passthruObj = [PSADT.Types.ProcessResult]::new(
                        $returnCode,
                        $(if (![System.String]::IsNullOrWhiteSpace($stdOut)) { $stdOut }),
                        $(if (![System.String]::IsNullOrWhiteSpace($stdErr)) { $stdErr })
                    )

                    # If we have an error in our process, throw it and let the catch block handle it.
                    if ($errorMessage)
                    {
                        $naerParams = @{
                            Exception = [System.ApplicationException]::new($errorMessage)
                            Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                            ErrorId = 'ProcessExitCodeError'
                            TargetObject = $passthruObj
                            RecommendedAction = "Please review the exit code with the vendor's documentation and try again."
                        }
                        throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                    }

                    # Update the session's last exit code with the value if externally called.
                    if ($adtSession -and $extInvoker)
                    {
                        $adtSession.SetExitCode($returnCode)
                    }

                    # If the passthru switch is specified, return the exit code and any output from process.
                    if ($PassThru)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'PassThru parameter specified, returning execution results object.'
                        $PSCmdlet.WriteObject($passthruObj)
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Set up parameters for Invoke-ADTFunctionErrorHandler.
            if ($null -ne $returnCode)
            {
                # Update the session's last exit code with the value if externally called.
                if ($adtSession -and $extInvoker -and ($OriginalErrorAction -notmatch '^(SilentlyContinue|Ignore)$'))
                {
                    $adtSession.SetExitCode($returnCode)
                }
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage $_.Exception.Message -DisableErrorResolving
            }
            else
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Error occurred while attempting to start the specified process."
            }

            # If the passthru switch is specified, return the exit code and any output from process.
            if ($PassThru)
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'PassThru parameter specified, returning execution results object.'
                $PSCmdlet.WriteObject($_.TargetObject)
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Start-ADTProcessAsUser
#
#-----------------------------------------------------------------------------

function Start-ADTProcessAsUser
{
    <#
    .SYNOPSIS
        Invokes a process in another user's session.

    .DESCRIPTION
        Invokes a process from SYSTEM in another user's session.

    .PARAMETER FilePath
        Path to the executable to invoke.

    .PARAMETER ArgumentList
        Arguments for the invoked executable.

    .PARAMETER WorkingDirectory
        The 'start-in' directory for the invoked executable.

    .PARAMETER HideWindow
        Specifies whether the window should be hidden or not.

    .PARAMETER ProcessCreationFlags
        One or more flags to control the process's invocation.

    .PARAMETER InheritEnvironmentVariables
        Specifies whether the process should inherit the user's environment state.

    .PARAMETER Wait
        Specifies whether to wait for the invoked excecutable to finish.

    .PARAMETER Username
        The username of the user's session to invoke the executable in.

    .PARAMETER SessionId
        The session ID of the user to invoke the executable in.

    .PARAMETER AllActiveUserSessions
        Specifies that the executable should be invoked in all active sessions.

    .PARAMETER UseLinkedAdminToken
        Specifies that an admin token (if available) should be used for the invocation.

    .PARAMETER SuccessExitCodes
        Specifies one or more exit codes that the function uses to consider the invocation successful.

    .PARAMETER ConsoleTimeoutInSeconds
        Specifies the timeout in seconds to wait for a console application to finish its task.

    .PARAMETER IsGuiApplication
        Indicates that the executed application is a GUI-based app, not a console-based app.

    .PARAMETER NoRedirectOutput
        Specifies that stdout/stderr output should not be redirected to file.

    .PARAMETER MergeStdErrAndStdOut
        Specifies that the stdout/stderr streams should be merged into a single output.

    .PARAMETER OutputDirectory
        Specifies the output directory for the redirected stdout/stderr streams.

    .PARAMETER NoTerminateOnTimeout
        Specifies that the process shouldn't terminate on timeout.

    .PARAMETER AdditionalEnvironmentVariables
        Specifies additional environment variables to inject into the user's session.

    .PARAMETER WaitOption
        Specifies the wait type to use when waiting for an invoked executable to finish.

    .PARAMETER SecureArgumentList
        Hides all parameters passed to the executable from the Toolkit log file.

    .PARAMETER PassThru
        If NoWait is not specified, returns an object with ExitCode, STDOut and STDErr output from the process. If NoWait is specified, returns an object with Id, Handle and ProcessName.

    .EXAMPLE
        Start-ADTProcessAsUser -FilePath "$($adtSession.DirFiles)\setup.exe" -ArgumentList '/S' -SuccessExitCodes 0, 500

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Threading.Tasks.Task[System.Int32]

        Returns a task object indicating the process's result.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding(DefaultParameterSetName = 'PrimaryActiveUserSession')]
    [OutputType([System.Threading.Tasks.Task[System.Int32]])]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $true, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $true, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $true, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.String]$FilePath,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$ArgumentList,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.String]$WorkingDirectory,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$HideWindow,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [PSADT.PInvoke.CREATE_PROCESS]$ProcessCreationFlags,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$InheritEnvironmentVariables,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$Wait,

        [Parameter(Mandatory = $true, ParameterSetName = 'Username')]
        [ValidateNotNullOrEmpty()]
        [System.String]$Username,

        [Parameter(Mandatory = $true, ParameterSetName = 'SessionId')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$SessionId,

        [Parameter(Mandatory = $true, ParameterSetName = 'AllActiveUserSessions')]
        [System.Management.Automation.SwitchParameter]$AllActiveUserSessions,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$UseLinkedAdminToken,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]$SuccessExitCodes,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.UInt32]$ConsoleTimeoutInSeconds,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$IsGuiApplication,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$NoRedirectOutput,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$MergeStdErrAndStdOut,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.String]$OutputDirectory,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$NoTerminateOnTimeout,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.IDictionary]$AdditionalEnvironmentVariables,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [ValidateNotNullOrEmpty()]
        [PSADT.ProcessEx.WaitType]$WaitOption,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false, ParameterSetName = 'Username')]
        [Parameter(Mandatory = $false, ParameterSetName = 'SessionId')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllActiveUserSessions')]
        [Parameter(Mandatory = $false, ParameterSetName = 'PrimaryActiveUserSession')]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        # Initialise function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Strip out parameters not destined for the C# code.
        $null = ('SecureArgumentList', 'PassThru').ForEach({
                if ($PSBoundParameters.ContainsKey($_))
                {
                    $PSBoundParameters.Remove($_)
                }
            })

        # If we're on the default parameter set, pass the right parameter through.
        if ($PSCmdlet.ParameterSetName.Equals('PrimaryActiveUserSession'))
        {
            $PSBoundParameters.Add('PrimaryActiveUserSession', [System.Management.Automation.SwitchParameter]$true)
        }
        elseif ($PSBoundParameters.ContainsKey('Username'))
        {
            if (!($userSessionId = & $Script:CommandTable.'Get-ADTLoggedOnUser' | & { process { if ($_ -and $_.NTAccount.EndsWith($Username, [System.StringComparison]::InvariantCultureIgnoreCase)) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1 -ExpandProperty SessionId))
            {
                $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Username -ProvidedValue $Username -ExceptionMessage 'An active session could not be found for the specified user.'))
            }
            $PSBoundParameters.Add('SessionId', ($SessionId = $userSessionId))
            $null = $PSBoundParameters.Remove('Username')
        }

        # Translate the environment variables into a dictionary. Using this type on the parameter is too hard on the caller.
        if ($PSBoundParameters.ContainsKey('AdditionalEnvironmentVariables'))
        {
            $AdditionalEnvironmentVariables = [System.Collections.Generic.Dictionary[System.String, System.String]]::new()
            $PSBoundParameters.AdditionalEnvironmentVariables.GetEnumerator() | & {
                process
                {
                    $AdditionalEnvironmentVariables.Add($_.Key, $_.Value)
                }
            }
            $PSBoundParameters.AdditionalEnvironmentVariables = $AdditionalEnvironmentVariables
        }

        # Translate switches that require negation for the LaunchOptions.
        $null = ('RedirectOutput', 'TerminateOnTimeout').Where({ $PSBoundParameters.ContainsKey("No$_") }).ForEach({
                $PSBoundParameters.$_ = !$PSBoundParameters."No$_"
                $PSBoundParameters.Remove("No$_")
            })

        # Unless explicitly provided, don't terminate on timeout.
        if (!$PSBoundParameters.ContainsKey('TerminateOnTimeout'))
        {
            $PSBoundParameters.TerminateOnTimeout = $false
        }

        # Translate the process flags into a list of flags. No idea why the backend is coded like this...
        if ($PSBoundParameters.ContainsKey('ProcessCreationFlags'))
        {
            $PSBoundParameters.ProcessCreationFlags = $PSBoundParameters.ProcessCreationFlags.ToString().Split(',').Trim()
        }
    }

    process
    {
        # Announce start.
        switch ($PSCmdlet.ParameterSetName)
        {
            Username
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoking [$FilePath$(if ($ArgumentList -and !$SecureArgumentList) { " $ArgumentList" })] as user [$Username]$(if ($Wait) { ", and waiting for invocation to finish" })."
                break
            }
            SessionId
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoking [$FilePath$(if ($ArgumentList -and !$SecureArgumentList) { " $ArgumentList" })] for session [$SessionId]$(if ($Wait) { ", and waiting for invocation to finish" })."
                break
            }
            AllActiveUserSessions
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoking [$FilePath$(if ($ArgumentList -and !$SecureArgumentList) { " $ArgumentList" })] for all active user sessions$(if ($Wait) { ", and waiting for all invocations to finish" })."
                break
            }
            PrimaryActiveUserSession
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invoking [$FilePath$(if ($ArgumentList -and !$SecureArgumentList) { " $ArgumentList" })] for the primary user session$(if ($Wait) { ", and waiting for invocation to finish" })."
                break
            }
        }

        # Create a new process object and invoke an execution.
        try
        {
            try
            {
                if (($result = ($process = [PSADT.ProcessEx.StartProcess]::new()).ExecuteAndMonitorAsync($PSBoundParameters)) -and $PassThru)
                {
                    return $result
                }
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
        finally
        {
            # Dispose of the process object to ensure things are cleaned up properly.
            $process.Dispose()
        }
    }

    end
    {
        # Finalise function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Start-ADTServiceAndDependencies
#
#-----------------------------------------------------------------------------

function Start-ADTServiceAndDependencies
{
    <#
    .SYNOPSIS
        Start a Windows service and its dependencies.

    .DESCRIPTION
        This function starts a specified Windows service and its dependencies. It provides options to skip starting dependent services, wait for a service to get out of a pending state, and return the service object.

    .PARAMETER Service
        Specify the name of the service.

    .PARAMETER SkipDependentServices
        Choose to skip checking for and starting dependent services. Default is: $false.

    .PARAMETER PendingStatusWait
        The amount of time to wait for a service to get out of a pending state before continuing. Default is 60 seconds.

    .PARAMETER PassThru
        Return the System.ServiceProcess.ServiceController service object.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.ServiceProcess.ServiceController

        Returns the service object.

    .EXAMPLE
        Start-ADTServiceAndDependencies -Service 'wuauserv'

        Starts the Windows Update service and its dependencies.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!$_.Name)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Service -ProvidedValue $_ -ExceptionMessage 'The specified service does not exist.'))
                }
                return !!$_
            })]
        [System.ServiceProcess.ServiceController]$Service,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipDependentServices,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$PendingStatusWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                & $Script:CommandTable.'Invoke-ADTServiceAndDependencyOperation' -Operation Start @PSBoundParameters
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to start the service [$($Service.Name)]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Stop-ADTServiceAndDependencies
#
#-----------------------------------------------------------------------------

function Stop-ADTServiceAndDependencies
{
    <#
    .SYNOPSIS
        Stop a Windows service and its dependencies.

    .DESCRIPTION
        This function stops a specified Windows service and its dependencies. It provides options to skip stopping dependent services, wait for a service to get out of a pending state, and return the service object.

    .PARAMETER Service
        Specify the name of the service.

    .PARAMETER SkipDependentServices
        Choose to skip checking for and stopping dependent services. Default is: $false.

    .PARAMETER PendingStatusWait
        The amount of time to wait for a service to get out of a pending state before continuing. Default is 60 seconds.

    .PARAMETER PassThru
        Return the System.ServiceProcess.ServiceController service object.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.ServiceProcess.ServiceController

        Returns the service object.

    .EXAMPLE
        Stop-ADTServiceAndDependencies -Service 'wuauserv'

        Stops the Windows Update service and its dependencies.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (!$_.Name)
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName Service -ProvidedValue $_ -ExceptionMessage 'The specified service does not exist.'))
                }
                return !!$_
            })]
        [System.ServiceProcess.ServiceController]$Service,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SkipDependentServices,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$PendingStatusWait,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                & $Script:CommandTable.'Invoke-ADTServiceAndDependencyOperation' -Operation Stop @PSBoundParameters
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to stop the service [$($Service.Name)]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTBattery
#
#-----------------------------------------------------------------------------

function Test-ADTBattery
{
    <#
    .SYNOPSIS
        Tests whether the local machine is running on AC power or not.

    .DESCRIPTION
        Tests whether the local machine is running on AC power and returns true/false. For detailed information, use the -PassThru option to get a hashtable containing various battery and power status properties.

    .PARAMETER PassThru
        Outputs a hashtable containing the following properties:
        - IsLaptop
        - IsUsingACPower
        - ACPowerLineStatus
        - BatteryChargeStatus
        - BatteryLifePercent
        - BatteryLifeRemaining
        - BatteryFullLifetime

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.BatteryInfo

        Returns a hashtable containing the following properties:
        - IsLaptop
        - IsUsingACPower
        - ACPowerLineStatus
        - BatteryChargeStatus
        - BatteryLifePercent
        - BatteryLifeRemaining
        - BatteryFullLifetime

    .EXAMPLE
        Test-ADTBattery

        Checks if the local machine is running on AC power and returns true or false.

    .EXAMPLE
        (Test-ADTBattery -PassThru).IsLaptop

        Returns true if the current system is a laptop, otherwise false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([PSADT.Types.BatteryInfo])]
    param
    (
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Checking if system is using AC power or if it is running on battery...'
        try
        {
            try
            {
                # Get the system power status. Indicates whether the system is using AC power or if the status is unknown. Possible values:
                # Offline : The system is not using AC power.
                # Online  : The system is using AC power.
                # Unknown : The power status of the system is unknown.
                $acPowerLineStatus = [System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus

                # Get the current battery charge status. Possible values: High, Low, Critical, Charging, NoSystemBattery, Unknown.
                $batteryChargeStatus = [System.Windows.Forms.SystemInformation]::PowerStatus.BatteryChargeStatus
                $invalidBattery = ($batteryChargeStatus -eq 'NoSystemBattery') -or ($batteryChargeStatus -eq 'Unknown')

                # Get the approximate amount, from 0.00 to 1.0, of full battery charge remaining.
                # This property can report 1.0 when the battery is damaged and Windows can't detect a battery.
                # Therefore, this property is only indicative of battery charge remaining if 'BatteryChargeStatus' property is not reporting 'NoSystemBattery' or 'Unknown'.
                $batteryLifePercent = [System.Windows.Forms.SystemInformation]::PowerStatus.BatteryLifePercent * !$invalidBattery

                # The reported approximate number of seconds of battery life remaining. It will report -1 if the remaining life is unknown because the system is on AC power.
                $batteryLifeRemainingSeconds = [System.Windows.Forms.SystemInformation]::PowerStatus.BatteryLifeRemaining

                # Get the manufacturer reported full charge lifetime of the primary battery power source in seconds.
                # The reported number of seconds of battery life available when the battery is fully charged, or -1 if it is unknown.
                # This will only be reported if the battery supports reporting this information. You will most likely get -1, indicating unknown.
                $batteryFullLifetimeSeconds = [System.Windows.Forms.SystemInformation]::PowerStatus.BatteryFullLifetime

                # Determine if the system is using AC power.
                $isUsingAcPower = switch ($acPowerLineStatus)
                {
                    Online
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'System is using AC power.'
                        $true
                        break
                    }
                    Offline
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'System is using battery power.'
                        $false
                        break
                    }
                    Unknown
                    {
                        if ($invalidBattery)
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "System power status is [$($acPowerLineStatus)] and battery charge status is [$batteryChargeStatus]. This is most likely due to a damaged battery so we will report system is using AC power."
                            $true
                        }
                        else
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message "System power status is [$($acPowerLineStatus)] and battery charge status is [$batteryChargeStatus]. Therefore, we will report system is using battery power."
                            $false
                        }
                        break
                    }
                }

                # Determine if the system is a laptop.
                $isLaptop = !$invalidBattery -and ((& $Script:CommandTable.'Get-CimInstance' -ClassName Win32_SystemEnclosure).ChassisTypes -match '^(9|10|14)$')

                # Return the object if we're passing through, otherwise just whether we're on AC.
                if ($PassThru)
                {
                    return [PSADT.Types.BatteryInfo]::new(
                        $acPowerLineStatus,
                        $batteryChargeStatus,
                        $batteryLifePercent,
                        $batteryLifeRemainingSeconds,
                        $batteryFullLifetimeSeconds,
                        $isUsingAcPower,
                        $isLaptop
                    )
                }
                return $isUsingAcPower
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTCallerIsAdmin
#
#-----------------------------------------------------------------------------

function Test-ADTCallerIsAdmin
{
    <#
    .SYNOPSIS
        Checks if the current user has administrative privileges.

    .DESCRIPTION
        This function checks if the current user is a member of the Administrators group. It returns a boolean value indicating whether the user has administrative privileges.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the current user is an administrator, otherwise $false.

    .EXAMPLE
        Test-ADTCallerIsAdmin

        Checks if the current user has administrative privileges and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    return [System.Security.Principal.WindowsPrincipal]::new([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTMicrophoneInUse
#
#-----------------------------------------------------------------------------

function Test-ADTMicrophoneInUse
{
    <#
    .SYNOPSIS
        Tests whether the device's microphone is in use.

    .DESCRIPTION
        Tests whether someone is using the microphone on their device. This could be within Teams, Zoom, a game, or any other app that uses a microphone.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the microphone is in use, otherwise returns $false.

    .EXAMPLE
        Test-ADTMicrophoneInUse

        Checks if the microphone is in use and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Checking whether the device's microphone is in use..."
        try
        {
            try
            {
                if (($microphoneInUse = [PSADT.Devices.Audio]::IsMicrophoneInUse()))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The device's microphone is currently in use."
                }
                else
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "The device's microphone is currently not in use."
                }
                return $microphoneInUse
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTModuleInitialized
#
#-----------------------------------------------------------------------------

function Test-ADTModuleInitialized
{
    <#
    .SYNOPSIS
        Checks if the ADT (PSAppDeployToolkit) module is initialized.

    .DESCRIPTION
        This function checks if the ADT (PSAppDeployToolkit) module is initialized by retrieving the module data and returning the initialization status.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the ADT module is initialized, otherwise $false.

    .EXAMPLE
        Test-ADTModuleInitialized

        Checks if the ADT module is initialized and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    return $Script:ADT.Initialized
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTMSUpdates
#
#-----------------------------------------------------------------------------

function Test-ADTMSUpdates
{
    <#
    .SYNOPSIS
        Test whether a Microsoft Windows update is installed.

    .DESCRIPTION
        This function checks if a specified Microsoft Windows update, identified by its KB number, is installed on the local machine. It first attempts to find the update using the Get-HotFix cmdlet and, if unsuccessful, uses a COM object to search the update history.

    .PARAMETER KbNumber
        KBNumber of the update.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the update is installed, otherwise returns $false.

    .EXAMPLE
        Test-ADTMSUpdates -KBNumber 'KB2549864'

        Checks if the Microsoft Update 'KB2549864' is installed and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = 'Enter the KB Number for the Microsoft Update')]
        [ValidateNotNullOrEmpty()]
        [System.String]$KbNumber
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Checking if Microsoft Update [$KbNumber] is installed."
        try
        {
            try
            {
                # Attempt to get the update via Get-HotFix first as it's cheap.
                if (!($kbFound = !!(& $Script:CommandTable.'Get-HotFix' -Id $KbNumber -ErrorAction Ignore)))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Unable to detect Windows update history via Get-Hotfix cmdlet. Trying via COM object.'
                    $updateSearcher = (& $Script:CommandTable.'New-Object' -ComObject Microsoft.Update.Session).CreateUpdateSearcher()
                    $updateSearcher.IncludePotentiallySupersededUpdates = $false
                    $updateSearcher.Online = $false
                    if (($updateHistoryCount = $updateSearcher.GetTotalHistoryCount()) -gt 0)
                    {
                        $kbFound = !!($updateSearcher.QueryHistory(0, $updateHistoryCount) | & { process { if (($_.Operation -ne 'Other') -and ($_.Title -match "\($KBNumber\)") -and ($_.Operation -eq 1) -and ($_.ResultCode -eq 2)) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1)
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Unable to detect Windows Update history via COM object.'
                        return
                    }
                }

                # Return result.
                if ($kbFound)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Microsoft Update [$KbNumber] is installed."
                    return $true
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Microsoft Update [$KbNumber] is not installed."
                return $false
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed discovering Microsoft Update [$kbNumber]."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTMutexAvailability
#
#-----------------------------------------------------------------------------

function Test-ADTMutexAvailability
{
    <#
    .SYNOPSIS
        Wait, up to a timeout value, to check if current thread is able to acquire an exclusive lock on a system mutex.

    .DESCRIPTION
        A mutex can be used to serialize applications and prevent multiple instances from being opened at the same time.

        Wait, up to a timeout (default is 1 millisecond), for the mutex to become available for an exclusive lock.

    .PARAMETER MutexName
        The name of the system mutex.

    .PARAMETER MutexWaitTime
        The number of milliseconds the current thread should wait to acquire an exclusive lock of a named mutex. Default is: 1 millisecond.

        A wait time of -1 milliseconds means to wait indefinitely. A wait time of zero does not acquire an exclusive lock but instead tests the state of the wait handle and returns immediately.

    .INPUTS
        None. You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean. Returns $true if the current thread acquires an exclusive lock on the named mutex, $false otherwise.

    .EXAMPLE
        Test-ADTMutexAvailability -MutexName 'Global\_MSIExecute' -MutexWaitTime 5000000

    .EXAMPLE
        Test-ADTMutexAvailability -MutexName 'Global\_MSIExecute' -MutexWaitTime (New-TimeSpan -Minutes 5)

    .EXAMPLE
        Test-ADTMutexAvailability -MutexName 'Global\_MSIExecute' -MutexWaitTime (New-TimeSpan -Seconds 60)

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        http://msdn.microsoft.com/en-us/library/aa372909(VS.85).asp

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateLength(1, 260)]
        [System.String]$MutexName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]$MutexWaitTime = [System.TimeSpan]::FromMilliseconds(1)
    )

    begin
    {
        # Initialize variables.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $WaitLogMsg = if ($MutexWaitTime.TotalMinutes -ge 1)
        {
            "$($MutexWaitTime.TotalMinutes) minute(s)"
        }
        elseif ($MutexWaitTime.TotalSeconds -ge 1)
        {
            "$($MutexWaitTime.TotalSeconds) second(s)"
        }
        else
        {
            "$($MutexWaitTime.Milliseconds) millisecond(s)"
        }
        $IsUnhandledException = $false
        $IsMutexFree = $false
        [System.Threading.Mutex]$OpenExistingMutex = $null
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Checking to see if mutex [$MutexName] is available. Wait up to [$WaitLogMsg] for the mutex to become available."
        try
        {
            # Open the specified named mutex, if it already exists, without acquiring an exclusive lock on it. If the system mutex does not exist, this method throws an exception instead of creating the system object.
            $OpenExistingMutex = [Threading.Mutex]::OpenExisting($MutexName)

            # Attempt to acquire an exclusive lock on the mutex. Use a Timespan to specify a timeout value after which no further attempt is made to acquire a lock on the mutex.
            $IsMutexFree = $OpenExistingMutex.WaitOne($MutexWaitTime, $false)
        }
        catch [Threading.WaitHandleCannotBeOpenedException]
        {
            # The named mutex does not exist.
            $IsMutexFree = $true
        }
        catch [ObjectDisposedException]
        {
            # Mutex was disposed between opening it and attempting to wait on it.
            $IsMutexFree = $true
        }
        catch [UnauthorizedAccessException]
        {
            # The named mutex exists, but the user does not have the security access required to use it.
            $IsMutexFree = $false
        }
        catch [Threading.AbandonedMutexException]
        {
            # The wait completed because a thread exited without releasing a mutex. This exception is thrown when one thread acquires a mutex object that another thread has abandoned by exiting without releasing it.
            $IsMutexFree = $true
        }
        catch
        {
            # Return $true, to signify that mutex is available, because function was unable to successfully complete a check due to an unhandled exception. Default is to err on the side of the mutex being available on a hard failure.
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Unable to check if mutex [$MutexName] is available due to an unhandled exception. Will default to return value of [$true].`n$(& $Script:CommandTable.'Resolve-ADTErrorRecord' -ErrorRecord $_)" -Severity 3
            $IsUnhandledException = $true
            $IsMutexFree = $true
        }
        finally
        {
            if ($IsMutexFree)
            {
                if (!$IsUnhandledException)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Mutex [$MutexName] is available for an exclusive lock."
                }
            }
            elseif (($MutexName -eq 'Global\_MSIExecute') -and ($msiInProgressCmdLine = & $Script:CommandTable.'Get-Process' -Name msiexec -ErrorAction Ignore | & { process { if ($_.CommandLine -match '\.msi') { $_.CommandLine.Trim() } } }))
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Mutex [$MutexName] is not available for an exclusive lock because the following MSI installation is in progress [$msiInProgressCmdLine]." -Severity 2
            }
            else
            {
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Mutex [$MutexName] is not available because another thread already has an exclusive lock on it."
            }

            if (($null -ne $OpenExistingMutex) -and $IsMutexFree)
            {
                # Release exclusive lock on the mutex.
                $null = $OpenExistingMutex.ReleaseMutex()
                $OpenExistingMutex.Close()
            }
        }
        return $IsMutexFree
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTNetworkConnection
#
#-----------------------------------------------------------------------------

function Test-ADTNetworkConnection
{
    <#
    .SYNOPSIS
        Tests for an active local network connection, excluding wireless and virtual network adapters.

    .DESCRIPTION
        Tests for an active local network connection, excluding wireless and virtual network adapters, by querying the Win32_NetworkAdapter WMI class. This function checks if any physical network adapter is in the 'Up' status.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if a wired network connection is detected, otherwise returns $false.

    .EXAMPLE
        Test-ADTNetworkConnection

        Checks if there is an active wired network connection and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Checking if system is using a wired network connection...'
        try
        {
            try
            {
                if (& $Script:CommandTable.'Get-NetAdapter' -Physical | & { process { if ($_.Status.Equals('Up')) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Wired network connection found.'
                    return $true
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Wired network connection not found.'
                return $false
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTOobeCompleted
#
#-----------------------------------------------------------------------------

function Test-ADTOobeCompleted
{
    <#
    .SYNOPSIS
        Checks if the device's Out-of-Box Experience (OOBE) has completed or not.

    .DESCRIPTION
        This function checks if the current device has completed the Out-of-Box Experience (OOBE).

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the device has proceeded past the OOBE, otherwise $false.

    .EXAMPLE
        Test-ADTOobeCompleted

        Checks if the device has completed the OOBE or not and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
    )

    begin
    {
        # Initialize function.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        # Return whether the OOBE is completed via an API call.
        try
        {
            try
            {
                return ([PSADT.Shared.Utility]::IsOOBEComplete())
            }
            catch
            {
                # Re-writing the ErrorRecord with Write-Error ensures the correct PositionMessage is used.
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            # Process the caught error, log it and throw depending on the specified ErrorAction.
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Error determining whether the OOBE has been completed or not."
        }
    }

    end
    {
        # Finalize function.
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTPowerPoint
#
#-----------------------------------------------------------------------------

function Test-ADTPowerPoint
{
    <#
    .SYNOPSIS
        Tests whether PowerPoint is running in either fullscreen slideshow mode or presentation mode.

    .DESCRIPTION
        Tests whether someone is presenting using PowerPoint in either fullscreen slideshow mode or presentation mode. This function checks if the PowerPoint process has a window with a title that begins with "PowerPoint Slide Show" or "PowerPoint-" for non-English language systems. There is a possibility of a false positive if the PowerPoint filename starts with "PowerPoint Slide Show". If the previous detection method does not detect PowerPoint in fullscreen mode, it checks if PowerPoint is in Presentation Mode (only works on Windows Vista or higher).

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if PowerPoint is running in either fullscreen slideshow mode or presentation mode, otherwise returns $false.

    .EXAMPLE
        Test-ADTPowerPoint

        Checks if PowerPoint is running in either fullscreen slideshow mode or presentation mode and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        This function can only execute detection logic if the process is in interactive mode.

        There is a possibility of a false positive if the PowerPoint filename starts with "PowerPoint Slide Show".

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $procName = 'POWERPNT'
        $presenting = 'Unknown'
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Checking if PowerPoint is in either fullscreen slideshow mode or presentation mode...'
        try
        {
            try
            {
                # Return early if we're not running PowerPoint or we can't interactively check.
                if (!($PowerPointProcess = & $Script:CommandTable.'Get-Process' -Name $procName -ErrorAction Ignore))
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'PowerPoint application is not running.'
                    return ($presenting = $false)
                }
                if (![System.Environment]::UserInteractive)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Unable to run check to see if PowerPoint is in fullscreen mode or Presentation Mode because current process is not interactive. Configure script to run in interactive mode in your deployment tool. If using SCCM Application Model, then make sure "Allow users to view and interact with the program installation" is selected. If using SCCM Package Model, then make sure "Allow users to interact with this program" is selected.' -Severity 2
                    return
                }

                # Check if "POWERPNT" process has a window with a title that begins with "PowerPoint Slide Show" or "Powerpoint-" for non-English language systems.
                # There is a possiblity of a false positive if the PowerPoint filename starts with "PowerPoint Slide Show".
                if (& $Script:CommandTable.'Get-ADTWindowTitle' -GetAllWindowTitles | & { process { if (($_.ParentProcess -eq $procName) -and ($_.WindowTitle -match '^PowerPoint(-| Slide Show)')) { return $_ } } } | & $Script:CommandTable.'Select-Object' -First 1)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Detected that PowerPoint process [$procName] has a window with a title that beings with [PowerPoint Slide Show] or [PowerPoint-]."
                    return ($presenting = $true)
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Detected that PowerPoint process [$procName] does not have a window with a title that beings with [PowerPoint Slide Show] or [PowerPoint-]."
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "PowerPoint process [$procName] has process ID(s) [$(($PowerPointProcessIDs = $PowerPointProcess.Id) -join ', ')]."

                # If previous detection method did not detect PowerPoint in fullscreen mode, then check if PowerPoint is in Presentation Mode (check only works on Windows Vista or higher).
                # Note: The below method does not detect PowerPoint presentation mode if the presentation is on a monitor that does not have current mouse input control.
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Detected user notification state [$(($UserNotificationState = [PSADT.GUI.UiAutomation]::GetUserNotificationState()))]."
                switch ($UserNotificationState)
                {
                    QUNS_PRESENTATION_MODE
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Detected that system is in [Presentation Mode].'
                        return ($presenting = $true)
                    }
                    QUNS_BUSY
                    {
                        if ($PowerPointProcessIDs -contains [PSADT.GUI.UiAutomation]::GetWindowThreadProcessId([PSADT.LibraryInterfaces.User32]::GetForegroundWindow()))
                        {
                            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Detected a fullscreen foreground window matches a PowerPoint process ID.'
                            return ($presenting = $true)
                        }
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Unable to find a fullscreen foreground window that matches a PowerPoint process ID.'
                        break
                    }
                }
                return ($presenting = $false)
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message "PowerPoint is running in fullscreen mode [$presenting]."
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTRegistryValue
#
#-----------------------------------------------------------------------------

function Test-ADTRegistryValue
{
    <#
    .SYNOPSIS
        Test if a registry value exists.

    .DESCRIPTION
        Checks a registry key path to see if it has a value with a given name. Can correctly handle cases where a value simply has an empty or null value.

    .PARAMETER Key
        Path of the registry key.

    .PARAMETER Name
        Specify the name of the value to check the existence of.

    .PARAMETER SID
        The security identifier (SID) for a user. Specifying this parameter will convert a HKEY_CURRENT_USER registry key to the HKEY_USERS\$SID format.

        Specify this parameter from the Invoke-ADTAllUsersRegistryAction function to read/edit HKCU registry settings for all users on the system.

    .PARAMETER Wow6432Node
        Specify this switch to check the 32-bit registry (Wow6432Node) on 64-bit systems.

    .INPUTS
        System.String

        Accepts a string value for the registry key path.

    .OUTPUTS
        System.Boolean

        Returns $true if the registry value exists, $false if it does not.

    .EXAMPLE
        Test-ADTRegistryValue -Key 'HKLM:SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations'

        Checks if the registry value 'PendingFileRenameOperations' exists under the specified key.

    .NOTES
        An active ADT session is NOT required to use this function.

        To test if a registry key exists, use the Test-Path function like so: Test-Path -LiteralPath $Key -PathType 'Container'

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Key,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Name,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]$SID,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$Wow6432Node
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # If the SID variable is specified, then convert all HKEY_CURRENT_USER key's to HKEY_USERS\$SID.
                $Key = if ($PSBoundParameters.ContainsKey('SID'))
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -Wow6432Node:$Wow6432Node -SID $SID
                }
                else
                {
                    & $Script:CommandTable.'Convert-ADTRegistryPath' -Key $Key -Wow6432Node:$Wow6432Node
                }

                # Test whether value exists or not.
                if ((& $Script:CommandTable.'Get-Item' -LiteralPath $Key -ErrorAction Ignore | & $Script:CommandTable.'Select-Object' -ExpandProperty Property -ErrorAction Ignore) -contains $Name)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Registry key value [$Key] [$Name] does exist."
                    return $true
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Registry key value [$Key] [$Name] does not exist."
                return $false
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTServiceExists
#
#-----------------------------------------------------------------------------

function Test-ADTServiceExists
{
    <#
    .SYNOPSIS
        Check to see if a service exists.

    .DESCRIPTION
        Check to see if a service exists. The UseCIM switch can be used in conjunction with PassThru to return WMI objects for PSADT v3.x compatibility, however, this method fails in Windows Sandbox.

    .PARAMETER Name
        Specify the name of the service.

        Note: Service name can be found by executing "Get-Service | Format-Table -AutoSize -Wrap" or by using the properties screen of a service in services.msc.

    .PARAMETER UseCIM
        Use CIM/WMI to check for the service. This is useful for compatibility with PSADT v3.x.

    .PARAMETER PassThru
        Return the WMI service object. To see all the properties use: Test-ADTServiceExists -Name 'spooler' -PassThru | Get-Member

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the service exists, otherwise returns $false.

    .EXAMPLE
        Test-ADTServiceExists -Name 'wuauserv'

        Checks if the service 'wuauserv' exists.

    .EXAMPLE
        Test-ADTServiceExists -Name 'testservice' -PassThru | Where-Object { $_ } | ForEach-Object { $_.Delete() }

        Checks if a service exists and then deletes it by using the -PassThru parameter.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = "This function is appropriately named and we don't need PSScriptAnalyzer telling us otherwise.")]
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Name,

        [Parameter(Mandatory = $false)]
        [Alias('UseWMI')]
        [System.Management.Automation.SwitchParameter]$UseCIM,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            try
            {
                # Access via CIM/WMI if specifically asked.
                if ($UseCIM)
                {
                    # If nothing is returned from Win32_Service, check Win32_BaseService.
                    if (!($ServiceObject = & $Script:CommandTable.'Get-CimInstance' -ClassName Win32_Service -Filter "Name = '$Name'"))
                    {
                        $ServiceObject = & $Script:CommandTable.'Get-CimInstance' -ClassName Win32_BaseService -Filter "Name = '$Name'"
                    }
                }
                else
                {
                    # If the result is empty, it means the provided service is invalid.
                    $ServiceObject = & $Script:CommandTable.'Get-Service' -Name $Name -ErrorAction Ignore
                }

                # Return early if null.
                if (!$ServiceObject)
                {
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Service [$Name] does not exist."
                    return $false
                }
                & $Script:CommandTable.'Write-ADTLogEntry' -Message "Service [$Name] exists."

                # Return the CIM object if passing through.
                if ($PassThru)
                {
                    return $ServiceObject
                }
                return $true
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed check to see if service [$Name] exists."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTSessionActive
#
#-----------------------------------------------------------------------------

function Test-ADTSessionActive
{
    <#
    .SYNOPSIS
        Checks if there is an active ADT session.

    .DESCRIPTION
        This function checks if there is an active ADT (App Deploy Toolkit) session by retrieving the module data and returning the count of active sessions.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if there is at least one active session, otherwise $false.

    .EXAMPLE
        Test-ADTSessionActive

        Checks if there is an active ADT session and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    return !!$Script:ADT.Sessions.Count
}


#-----------------------------------------------------------------------------
#
# MARK: Test-ADTUserIsBusy
#
#-----------------------------------------------------------------------------

function Test-ADTUserIsBusy
{
    <#
    .SYNOPSIS
        Tests whether the device's microphone is in use, the user has manually turned on presentation mode, or PowerPoint is running in either fullscreen slideshow mode or presentation mode.

    .DESCRIPTION
        Tests whether the device's microphone is in use, the user has manually turned on presentation mode, or PowerPoint is running in either fullscreen slideshow mode or presentation mode.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        System.Boolean

        Returns $true if the device's microphone is in use, the user has manually turned on presentation mode, or PowerPoint is running in either fullscreen slideshow mode or presentation mode, otherwise $false.

    .EXAMPLE
        Test-ADTUserIsBusy

        Tests whether the device's microphone is in use, the user has manually turned on presentation mode, or PowerPoint is running in either fullscreen slideshow mode or presentation mode, and returns true or false.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
    )

    try
    {
        return ((& $Script:CommandTable.'Test-ADTMicrophoneInUse') -or (& $Script:CommandTable.'Get-ADTPresentationSettingsEnabledUsers') -or (& $Script:CommandTable.'Test-ADTPowerPoint'))
    }
    catch
    {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Unblock-ADTAppExecution
#
#-----------------------------------------------------------------------------

function Unblock-ADTAppExecution
{
    <#
    .SYNOPSIS
        Unblocks the execution of applications performed by the Block-ADTAppExecution function.

    .DESCRIPTION
        This function is called by the Close-ADTSession function or when the script itself is called with the parameters -CleanupBlockedApps. It undoes the actions performed by Block-ADTAppExecution, allowing previously blocked applications to execute.

    .PARAMETER Tasks
        Specify the scheduled tasks to unblock.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not generate any output.

    .EXAMPLE
        Unblock-ADTAppExecution

        Unblocks the execution of applications that were previously blocked by Block-ADTAppExecution.

    .NOTES
        An active ADT session is NOT required to use this function.

        It is used when the -BlockExecution parameter is specified with the Show-ADTInstallationWelcome function to undo the actions performed by Block-ADTAppExecution.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Management.Infrastructure.CimInstance[]]$Tasks = (& $Script:CommandTable.'Get-ScheduledTask' -TaskName "$($MyInvocation.MyCommand.Module.Name)_*_BlockedApps" -ErrorAction Ignore)
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $uaaeiParams = @{}; if ($Tasks) { $uaaeiParams.Add('Tasks', $Tasks) }
    }

    process
    {
        # Bypass if no admin rights.
        if (!(& $Script:CommandTable.'Test-ADTCallerIsAdmin'))
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message "Bypassing Function [$($MyInvocation.MyCommand.Name)], because [User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)] is not admin."
            return
        }

        # Clean up blocked apps using our backend worker.
        try
        {
            try
            {
                & $Script:CommandTable.'Unblock-ADTAppExecutionInternal' @uaaeiParams -Verbose 4>&1 | & $Script:CommandTable.'Write-ADTLogEntry'
                & $Script:CommandTable.'Remove-ADTSessionFinishingCallback' -Callback $MyInvocation.MyCommand
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Uninstall-ADTApplication
#
#-----------------------------------------------------------------------------

function Uninstall-ADTApplication
{
    <#
    .SYNOPSIS
        Removes one or more applications specified by name, filter script, or InstalledApplication object from Get-ADTApplication.

    .DESCRIPTION
        Removes one or more applications specified by name, filter script, or InstalledApplication object from Get-ADTApplication.

        Enumerates the registry for installed applications via Get-ADTApplication, matching the specified application name and uninstalls that application using its uninstall string, with the ability to specify additional uninstall parameters also.

    .PARAMETER InstalledApplication
        Specifies the [PSADT.Types.InstalledApplication] object to remove. This parameter is typically used when piping Get-ADTApplication to this function.

    .PARAMETER Name
        The name of the application to retrieve information for. Performs a contains match on the application display name by default.

    .PARAMETER NameMatch
        Specifies the type of match to perform on the application name. Valid values are 'Contains', 'Exact', 'Wildcard', and 'Regex'. The default value is 'Contains'.

    .PARAMETER ProductCode
        The product code of the application to retrieve information for.

    .PARAMETER ApplicationType
        Specifies the type of application to remove. Valid values are 'All', 'MSI', and 'EXE'. The default value is 'All'.

    .PARAMETER IncludeUpdatesAndHotfixes
        Include matches against updates and hotfixes in results.

    .PARAMETER FilterScript
        A script used to filter the results as they're processed.

    .PARAMETER ArgumentList
        Overrides the default MSI parameters specified in the config.psd1 file, or the parameters found in QuietUninstallString/UninstallString for EXE applications.

    .PARAMETER AdditionalArgumentList
        Adds to the default parameters specified in the config.psd1 file, or the parameters found in QuietUninstallString/UninstallString for EXE applications.

    .PARAMETER SecureArgumentList
        Hides all parameters passed to the executable from the Toolkit log file.

    .PARAMETER LoggingOptions
        Overrides the default MSI logging options specified in the config.psd1 file. Default options are: "/L*v".

    .PARAMETER LogFileName
        Overrides the default log file name for MSI applications. The default log file name is generated from the MSI file name. If LogFileName does not end in .log, it will be automatically appended.

        For uninstallations, by default the product code is resolved to the DisplayName and version of the application.

    .PARAMETER PassThru
        Returns ExitCode, STDOut, and STDErr output from the process.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        PSADT.Types.ProcessResult

        Returns an object with the results of the installation if -PassThru is specified.
        - ExitCode
        - StdOut
        - StdErr

    .EXAMPLE
        Uninstall-ADTApplication -Name 'Acrobat' -ApplicationType 'MSI' -FilterScript { $_.Publisher -match 'Adobe' }

        Removes all MSI applications that contain the name 'Acrobat' in the DisplayName and 'Adobe' in the Publisher name.

    .EXAMPLE
        Uninstall-ADTApplication -Name 'Java' -FilterScript {$_.Publisher -eq 'Oracle Corporation' -and $_.Is64BitApplication -eq $true -and $_.DisplayVersion -notlike '8.*'}

        Removes all MSI applications that contain the name 'Java' in the DisplayName, with Publisher as 'Oracle Corporation', are 64-bit, and not version 8.x.

    .EXAMPLE
        Uninstall-ADTApplication -FilterScript {$_.DisplayName -match '^Vim\s'} -Verbose -ApplicationType EXE -ArgumentList '/S'

        Remove all EXE applications starting with the name 'Vim' followed by a space, using the '/S' parameter.

    .NOTES
        An active ADT session is NOT required to use this function.

        More reading on how to create filterscripts https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-5.1#description

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'NameMatch', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ApplicationType', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'IncludeUpdatesAndHotfixes', Justification = "This parameter is used within delegates that PSScriptAnalyzer has no visibility of. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1472 for more details.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'LoggingOptions', Justification = "This parameter is used/retrieved via Get-ADTBoundParametersAndDefaultValues, which is too advanced for PSScriptAnalyzer to comprehend.")]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'LogFileName', Justification = "This parameter is used/retrieved via Get-ADTBoundParametersAndDefaultValues, which is too advanced for PSScriptAnalyzer to comprehend.")]
    [CmdletBinding()]
    [OutputType([PSADT.Types.ProcessResult])]
    [OutputType([PSADT.Types.ProcessInfo])]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'InstalledApplication', ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [PSADT.Types.InstalledApplication[]]$InstalledApplication,

        [Parameter(Mandatory = $false, ParameterSetName = 'Search')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]$Name,

        [Parameter(Mandatory = $false, ParameterSetName = 'Search')]
        [ValidateSet('Contains', 'Exact', 'Wildcard', 'Regex')]
        [System.String]$NameMatch = 'Contains',

        [Parameter(Mandatory = $false, ParameterSetName = 'Search')]
        [ValidateNotNullOrEmpty()]
        [System.Guid[]]$ProductCode,

        [Parameter(Mandatory = $false, ParameterSetName = 'Search')]
        [ValidateSet('All', 'MSI', 'EXE')]
        [System.String]$ApplicationType = 'All',

        [Parameter(Mandatory = $false, ParameterSetName = 'Search')]
        [System.Management.Automation.SwitchParameter]$IncludeUpdatesAndHotfixes,

        [Parameter(Mandatory = $false, ParameterSetName = 'Search', Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]$FilterScript,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$AdditionalArgumentList,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$SecureArgumentList,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LoggingOptions,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogFileName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.SwitchParameter]$PassThru
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue

        if ($PSCmdlet.ParameterSetName -ne 'InstalledApplication')
        {
            if (!($PSBoundParameters.Keys -match '^(Name|ProductCode|FilterScript)$'))
            {
                $naerParams = @{
                    Exception = [System.ArgumentNullException]::new('Either Name, ProductCode or FilterScript are required if not using pipeline.')
                    Category = [System.Management.Automation.ErrorCategory]::InvalidArgument
                    ErrorId = 'NullParameterValue'
                    RecommendedAction = "Review the supplied parameter values and try again."
                }
                $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTErrorRecord' @naerParams))
            }

            # Build the hashtable with the options that will be passed to Get-ADTApplication using splatting
            $gaiaParams = & $Script:CommandTable.'Get-ADTBoundParametersAndDefaultValues' -Invocation $MyInvocation -ParameterSetName $PSCmdlet.ParameterSetName -Exclude ArgumentList, AdditionalArgumentList, LoggingOptions, LogFileName, PassThru, SecureArgumentList
            if (($installedApps = & $Script:CommandTable.'Get-ADTApplication' @gaiaParams))
            {
                $InstalledApplication = $installedApps
            }
        }

        # Build the hashtable with the options that will be passed to Start-ADTMsiProcess using splatting
        $sampParams = & $Script:CommandTable.'Get-ADTBoundParametersAndDefaultValues' -Invocation $MyInvocation -ParameterSetName $PSCmdlet.ParameterSetName -Exclude InstalledApplication, Name, NameMatch, ProductCode, FilterScript, ApplicationType
        $sampParams.Action = 'Uninstall'

        # Build the hashtable with the options that will be passed to Start-ADTProcess using splatting.
        $sapParams = @{
            SecureArgumentList = $SecureArgumentList
            WaitForMsiExec = $true
            CreateNoWindow = $true
            PassThru = $PassThru
        }
    }

    process
    {
        if (!$InstalledApplication)
        {
            & $Script:CommandTable.'Write-ADTLogEntry' -Message 'No applications found for removal.'
            return
        }

        foreach ($removeApplication in $InstalledApplication)
        {
            try
            {
                if ($removeApplication.WindowsInstaller)
                {
                    if (!$removeApplication.ProductCode)
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "No ProductCode found for MSI application [$($removeApplication.DisplayName) $($removeApplication.DisplayVersion)]. Skipping removal."
                        continue
                    }
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing MSI application [$($removeApplication.DisplayName) $($removeApplication.DisplayVersion)] with ProductCode [$($removeApplication.ProductCode.ToString('B'))]."
                    try
                    {
                        if ($sampParams.ContainsKey('FilePath'))
                        {
                            $null = $sampParams.Remove('FilePath')
                        }
                        $removeApplication | & $Script:CommandTable.'Start-ADTMsiProcess' @sampParams
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                    }
                }
                else
                {
                    $uninstallString = if (![System.String]::IsNullOrWhiteSpace($removeApplication.QuietUninstallString))
                    {
                        $removeApplication.QuietUninstallString
                    }
                    elseif (![System.String]::IsNullOrWhiteSpace($removeApplication.UninstallString))
                    {
                        $removeApplication.UninstallString
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "No UninstallString found for EXE application [$($removeApplication.DisplayName) $($removeApplication.DisplayVersion)]. Skipping removal."
                        continue
                    }

                    $invalidFileNameChars = [System.Text.RegularExpressions.Regex]::Escape([System.String]::Join($null, [System.IO.Path]::GetInvalidFileNameChars()))
                    $invalidPathChars = [System.Text.RegularExpressions.Regex]::Escape([System.String]::Join($null, [System.IO.Path]::GetInvalidPathChars()))

                    if ($uninstallString -match "^`"?([^$invalidFileNameChars\s]+(?=\s|$)|[^$invalidPathChars]+?\.(?:exe|cmd|bat|vbs))`"?(?:\s(.*))?$")
                    {
                        $sapParams.FilePath = [System.Environment]::ExpandEnvironmentVariables($matches[1])
                        if (![System.IO.File]::Exists($sapParams.FilePath) -and ($commandPath = & $Script:CommandTable.'Get-Command' -Name $sapParams.FilePath -ErrorAction Ignore))
                        {
                            $sapParams.FilePath = $commandPath.Source
                        }
                        $uninstallStringParams = if ($matches.Count -gt 2)
                        {
                            [System.Environment]::ExpandEnvironmentVariables($matches[2].Trim())
                        }
                    }
                    else
                    {
                        & $Script:CommandTable.'Write-ADTLogEntry' -Message "Invalid UninstallString [$uninstallString] found for EXE application [$($removeApplication.DisplayName) $($removeApplication.DisplayVersion)]. Skipping removal."
                        continue
                    }

                    if (![System.String]::IsNullOrWhiteSpace($ArgumentList))
                    {
                        $sapParams.ArgumentList = $ArgumentList
                    }
                    elseif (![System.String]::IsNullOrWhiteSpace($uninstallStringParams))
                    {
                        $sapParams.ArgumentList = $uninstallStringParams
                    }
                    else
                    {
                        $sapParams.Remove('ArgumentList')
                    }
                    if ($AdditionalArgumentList)
                    {
                        if ($sapParams.ContainsKey('ArgumentList'))
                        {
                            $sapParams.ArgumentList += " $([System.String]::Join(' ', $AdditionalArgumentList))"
                        }
                        else
                        {
                            $sapParams.ArgumentList = $AdditionalArgumentList
                        }
                    }

                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "Removing EXE application [$($removeApplication.DisplayName) $($removeApplication.DisplayVersion)]."
                    try
                    {
                        & $Script:CommandTable.'Start-ADTProcess' @sapParams
                    }
                    catch
                    {
                        & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                    }
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Unregister-ADTDll
#
#-----------------------------------------------------------------------------

function Unregister-ADTDll
{
    <#
    .SYNOPSIS
        Unregister a DLL file.

    .DESCRIPTION
        Unregister a DLL file using regsvr32.exe. This function takes the path to the DLL file and attempts to unregister it using the regsvr32.exe utility.

    .PARAMETER FilePath
        Path to the DLL file.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return objects.

    .EXAMPLE
        Unregister-ADTDll -FilePath "C:\Test\DcTLSFileToDMSComp.dll"

        Unregisters the specified DLL file.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if (![System.IO.File]::Exists($_))
                {
                    $PSCmdlet.ThrowTerminatingError((& $Script:CommandTable.'New-ADTValidateScriptErrorRecord' -ParameterName FilePath -ProvidedValue $_ -ExceptionMessage 'The specified file does not exist.'))
                }
                return ![System.String]::IsNullOrWhiteSpace($_)
            })]
        [System.String]$FilePath
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            & $Script:CommandTable.'Invoke-ADTRegSvr32' @PSBoundParameters -Action Unregister
        }
        catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Update-ADTDesktop
#
#-----------------------------------------------------------------------------

function Update-ADTDesktop
{
    <#
    .SYNOPSIS
        Refresh the Windows Explorer Shell, which causes the desktop icons and the environment variables to be reloaded.

    .DESCRIPTION
        This function refreshes the Windows Explorer Shell, causing the desktop icons and environment variables to be reloaded. This can be useful after making changes that affect the desktop or environment variables, ensuring that the changes are reflected immediately.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Update-ADTDesktop

        Refreshes the Windows Explorer Shell, reloading the desktop icons and environment variables.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    begin
    {
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Refreshing the Desktop and the Windows Explorer environment process block.'
        try
        {
            try
            {
                [PSADT.GUI.Explorer]::RefreshDesktopAndEnvironmentVariables()
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to refresh the Desktop and the Windows Explorer environment process block."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Update-ADTEnvironmentPsProvider
#
#-----------------------------------------------------------------------------

function Update-ADTEnvironmentPsProvider
{
    <#
    .SYNOPSIS
        Updates the environment variables for the current PowerShell session with any environment variable changes that may have occurred during script execution.

    .DESCRIPTION
        Environment variable changes that take place during script execution are not visible to the current PowerShell session.
        Use this function to refresh the current PowerShell session with all environment variable settings.

    .PARAMETER LoadLoggedOnUserEnvironmentVariables
        If script is running in SYSTEM context, this option allows loading environment variables from the active console user. If no console user exists but users are logged in, such as on terminal servers, then the first logged-in non-console user.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Update-ADTEnvironmentPsProvider

        Refreshes the current PowerShell session with all environment variable settings.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$LoadLoggedOnUserEnvironmentVariables
    )

    begin
    {
        # Perform initial setup.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

        # Determine the user SID to base things off of.
        $userSid = if ($LoadLoggedOnUserEnvironmentVariables -and ($runAsActiveUser = & $Script:CommandTable.'Get-ADTRunAsActiveUser'))
        {
            $runAsActiveUser.SID
        }
        else
        {
            [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        }
    }

    process
    {
        & $Script:CommandTable.'Write-ADTLogEntry' -Message 'Refreshing the environment variables for this PowerShell session.'
        try
        {
            try
            {
                # Update all session environment variables. Ordering is important here: user variables comes second so that we can override system variables.
                & $Script:CommandTable.'Get-ItemProperty' -LiteralPath 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment', "Microsoft.PowerShell.Core\Registry::HKEY_USERS\$userSid\Environment" | & {
                    process
                    {
                        $_.PSObject.Properties | & {
                            process
                            {
                                if ($_.Name -notmatch '^PS((Parent)?Path|ChildName|Provider)$')
                                {
                                    [System.Environment]::SetEnvironmentVariable($_.Name, $_.Value)
                                }
                            }
                        }
                    }
                }

                # Set PATH environment variable separately because it is a combination of the user and machine environment variables.
                [System.Environment]::SetEnvironmentVariable('PATH', [System.String]::Join(';', (([System.Environment]::GetEnvironmentVariable('PATH', 'Machine'), [System.Environment]::GetEnvironmentVariable('PATH', 'User')).Split(';').Trim() | & { process { if ($_) { return $_ } } } | & $Script:CommandTable.'Select-Object' -Unique)))
            }
            catch
            {
                & $Script:CommandTable.'Write-Error' -ErrorRecord $_
            }
        }
        catch
        {
            & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_ -LogMessage "Failed to refresh the environment variables for this PowerShell session."
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Update-ADTGroupPolicy
#
#-----------------------------------------------------------------------------

function Update-ADTGroupPolicy
{
    <#
    .SYNOPSIS
        Performs a gpupdate command to refresh Group Policies on the local machine.

    .DESCRIPTION
        This function performs a gpupdate command to refresh Group Policies on the local machine. It updates both Computer and User policies by forcing a refresh using the gpupdate.exe utility.

    .INPUTS
        None

        You cannot pipe objects to this function.

    .OUTPUTS
        None

        This function does not return any objects.

    .EXAMPLE
        Update-ADTGroupPolicy

        Performs a gpupdate command to refresh Group Policies on the local machine.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    param
    (
    )

    begin
    {
        # Make this function continue on error.
        & $Script:CommandTable.'Initialize-ADTFunction' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorAction SilentlyContinue
    }

    process
    {
        # Handle each target separately so we can report on it.
        foreach ($target in ('Computer', 'User'))
        {
            try
            {
                try
                {
                    # Invoke gpupdate.exe and cache the results. An exit code of 0 is considered successful.
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message "$(($msg = "Updating Group Policies for the $target"))."
                    $gpUpdateResult = & "$([System.Environment]::SystemDirectory)\cmd.exe" /c "echo N | gpupdate.exe /Target:$target /Force" 2>&1
                    if (!$Global:LASTEXITCODE)
                    {
                        return
                    }

                    # If we're here, we had a bad exit code.
                    & $Script:CommandTable.'Write-ADTLogEntry' -Message ($msg = "$msg failed with exit code [$Global:LASTEXITCODE].") -Severity 3
                    $naerParams = @{
                        Exception = [System.Runtime.InteropServices.ExternalException]::new($msg, $Global:LASTEXITCODE)
                        Category = [System.Management.Automation.ErrorCategory]::InvalidResult
                        ErrorId = 'GpUpdateFailure'
                        TargetObject = $gpUpdateResult
                        RecommendedAction = "Please review the result in this error's TargetObject property and try again."
                    }
                    throw (& $Script:CommandTable.'New-ADTErrorRecord' @naerParams)
                }
                catch
                {
                    & $Script:CommandTable.'Write-Error' -ErrorRecord $_
                }
            }
            catch
            {
                & $Script:CommandTable.'Invoke-ADTFunctionErrorHandler' -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
            }
        }
    }

    end
    {
        & $Script:CommandTable.'Complete-ADTFunction' -Cmdlet $PSCmdlet
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Write-ADTLogEntry
#
#-----------------------------------------------------------------------------

function Write-ADTLogEntry
{
    <#
    .SYNOPSIS
        Write messages to a log file in CMTrace.exe compatible format or Legacy text file format.

    .DESCRIPTION
        Write messages to a log file in CMTrace.exe compatible format or Legacy text file format and optionally display in the console. This function supports different severity levels and can be used to log debug messages if required.

    .PARAMETER Message
        The message to write to the log file or output to the console.

    .PARAMETER Severity
        Defines message type. When writing to console or CMTrace.exe log format, it allows highlighting of message type.
        Options: 0 = Success (highlighted in green), 1 = Information (default), 2 = Warning (highlighted in yellow), 3 = Error (highlighted in red)

    .PARAMETER Source
        The source of the message being logged.

    .PARAMETER ScriptSection
        The heading for the portion of the script that is being executed. Default is: "$($adtSession.InstallPhase)".

    .PARAMETER LogType
        Choose whether to write a CMTrace.exe compatible log file or a Legacy text log file.

    .PARAMETER LogFileDirectory
        Set the directory where the log file will be saved.

    .PARAMETER LogFileName
        Set the name of the log file.

    .PARAMETER PassThru
        Return the message that was passed to the function.

    .PARAMETER DebugMessage
        Specifies that the message is a debug message. Debug messages only get logged if -LogDebugMessage is set to $true.

    .INPUTS
        System.String

        The message to write to the log file or output to the console.

    .OUTPUTS
        System.String[]

        This function returns the provided output if -PassThru is specified.

    .EXAMPLE
        Write-ADTLogEntry -Message "Installing patch MS15-031" -Source 'Add-Patch'

        Writes a log entry indicating that patch MS15-031 is being installed.

    .EXAMPLE
        Write-ADTLogEntry -Message "Script is running on Windows 11" -Source 'Test-ValidOS'

        Writes a log entry indicating that the script is running on Windows 11.

    .NOTES
        An active ADT session is NOT required to use this function.

        Tags: psadt
        Website: https://psappdeploytoolkit.com
        Copyright: (C) 2024 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).
        License: https://opensource.org/license/lgpl-3-0

    .LINK
        https://psappdeploytoolkit.com
    #>

    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.StringCollection])]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [System.String[]]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 3)]
        [System.Nullable[System.UInt32]]$Severity,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$Source,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$ScriptSection,

        [Parameter(Mandatory = $false)]
        [ValidateSet('CMTrace', 'Legacy')]
        [System.String]$LogType,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogFileDirectory,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$LogFileName,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$PassThru,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$DebugMessage
    )

    begin
    {
        # Get the caller's preference values and set them within this scope.
        & $Script:CommandTable.'Set-ADTPreferenceVariables' -SessionState $ExecutionContext.SessionState

        # Set up collector for piped in messages.
        $messages = [System.Collections.Specialized.StringCollection]::new()
    }

    process
    {
        # Return early if the InformationPreference is silent.
        if (($Severity -le 1) -and ($InformationPreference -match '^(SilentlyContinue|Ignore)$'))
        {
            return
        }

        # Add all non-null messages to the collector.
        $null = $Message | & {
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
        # Return early if we have no messages to write out.
        if (!$messages.Count)
        {
            return
        }

        # If we don't have an active session, write the message to the verbose stream (4).
        if (& $Script:CommandTable.'Test-ADTSessionActive')
        {
            (& $Script:CommandTable.'Get-ADTSession').WriteLogEntry($messages, $Severity, $Source, $ScriptSection, $null, $DebugMessage, $LogType, $LogFileDirectory, $LogFileName)
        }
        elseif (!$DebugMessage)
        {
            if ([System.String]::IsNullOrWhiteSpace($Source))
            {
                $Source = [PSADT.Module.DeploymentSession]::GetLogEntryCaller([System.Management.Automation.CallStackFrame[]](& $Script:CommandTable.'Get-PSCallStack')).Command
            }
            $messages -replace '^', "[$([System.DateTime]::Now.ToString('O'))] [$Source] :: " | & $Script:CommandTable.'Write-Verbose'
        }

        # Return the provided message if PassThru is true.
        if ($PassThru)
        {
            return $messages
        }
    }
}


#-----------------------------------------------------------------------------
#
# MARK: Module Constants and Function Exports
#
#-----------------------------------------------------------------------------

# Set all functions as read-only, export all public definitions and finalise the CommandTable.
& $Script:CommandTable.'Set-Item' -LiteralPath $FunctionPaths -Options ReadOnly
& $Script:CommandTable.'Get-Item' -LiteralPath $FunctionPaths | & { process { $CommandTable.Add($_.Name, $_) } }
& $Script:CommandTable.'New-Variable' -Name CommandTable -Value ([System.Collections.ObjectModel.ReadOnlyDictionary[System.String, System.Management.Automation.CommandInfo]]::new($CommandTable)) -Option Constant -Force -Confirm:$false
& $Script:CommandTable.'Export-ModuleMember' -Function $Module.Manifest.FunctionsToExport

# Define object for holding all PSADT variables.
& $Script:CommandTable.'New-Variable' -Name ADT -Option Constant -Value ([pscustomobject]@{
        Callbacks = [pscustomobject]@{
            Starting = [System.Collections.Generic.List[System.Management.Automation.CommandInfo]]::new()
            Opening = [System.Collections.Generic.List[System.Management.Automation.CommandInfo]]::new()
            Closing = [System.Collections.Generic.List[System.Management.Automation.CommandInfo]]::new()
            Finishing = [System.Collections.Generic.List[System.Management.Automation.CommandInfo]]::new()
        }
        Directories = [pscustomobject]@{
            Defaults = ([ordered]@{
                    Script = "$PSScriptRoot"
                    Config = "$PSScriptRoot\Config"
                    Strings = "$PSScriptRoot\Strings"
                }).AsReadOnly()
            Script = $null
            Config = $null
            Strings = $null
        }
        Durations = [pscustomobject]@{
            ModuleImport = $null
            ModuleInit = $null
        }
        Sessions = [System.Collections.Generic.List[PSADT.Module.DeploymentSession]]::new()
        SessionState = $ExecutionContext.SessionState
        TerminalServerMode = $false
        Environment = $null
        Language = $null
        Config = $null
        Strings = $null
        LastExitCode = 0
        Initialized = $false
    })

# Define object for holding all dialog window variables.
& $Script:CommandTable.'New-Variable' -Name Dialogs -Option Constant -Value ([ordered]@{
        Box = ([ordered]@{
                Buttons = ([ordered]@{
                        OK = 0
                        OKCancel = 1
                        AbortRetryIgnore = 2
                        YesNoCancel = 3
                        YesNo = 4
                        RetryCancel = 5
                        CancelTryAgainContinue = 6
                    }).AsReadOnly()
                Icons = ([ordered]@{
                        None = 0
                        Stop = 16
                        Question = 32
                        Exclamation = 48
                        Information = 64
                    }).AsReadOnly()
                DefaultButtons = ([ordered]@{
                        First = 0
                        Second = 256
                        Third = 512
                    }).AsReadOnly()
            }).AsReadOnly()
        Classic = [pscustomobject]@{
            ProgressWindow = [pscustomobject]@{
                SyncHash = [System.Collections.Hashtable]::Synchronized(@{})
                XamlCode = $null
                PowerShell = $null
                Invocation = $null
                Running = $false
            }
            Assets = [pscustomobject]@{
                Icon = $null
                Logo = $null
                Banner = $null
            }
            Font = [System.Drawing.SystemFonts]::MessageBoxFont
            BannerHeight = 0
            Width = 450
        }
        Fluent = [pscustomobject]@{
            ProgressWindow = [pscustomobject]@{
                Running = $false
            }
        }
    }).AsReadOnly()

# Registry path transformation constants used within Convert-ADTRegistryPath.
& $Script:CommandTable.'New-Variable' -Name Registry -Option Constant -Value ([ordered]@{
        PathMatches = [System.Collections.ObjectModel.ReadOnlyCollection[System.String]]$(
            ':\\'
            ':'
            '\\'
        )
        PathReplacements = ([ordered]@{
                '^HKLM' = 'HKEY_LOCAL_MACHINE\'
                '^HKCR' = 'HKEY_CLASSES_ROOT\'
                '^HKCU' = 'HKEY_CURRENT_USER\'
                '^HKU' = 'HKEY_USERS\'
                '^HKCC' = 'HKEY_CURRENT_CONFIG\'
                '^HKPD' = 'HKEY_PERFORMANCE_DATA\'
            }).AsReadOnly()
        WOW64Replacements = ([ordered]@{
                '^(HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\|HKEY_CURRENT_USER\\SOFTWARE\\Classes\\|HKEY_CLASSES_ROOT\\)(AppID\\|CLSID\\|DirectShow\\|Interface\\|Media Type\\|MediaFoundation\\|PROTOCOLS\\|TypeLib\\)' = '$1Wow6432Node\$2'
                '^HKEY_LOCAL_MACHINE\\SOFTWARE\\' = 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\'
                '^HKEY_LOCAL_MACHINE\\SOFTWARE$' = 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node'
                '^HKEY_CURRENT_USER\\Software\\Microsoft\\Active Setup\\Installed Components\\' = 'HKEY_CURRENT_USER\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\'
            }).AsReadOnly()
    }).AsReadOnly()

# Lookup table for preference variables and their associated CommonParameter name.
& $Script:CommandTable.'New-Variable' -Name PreferenceVariableTable -Option Constant -Value ([ordered]@{
        'InformationAction' = 'InformationPreference'
        'ProgressAction' = 'ProgressPreference'
        'WarningAction' = 'WarningPreference'
        'Confirm' = 'ConfirmPreference'
        'Verbose' = 'VerbosePreference'
        'WhatIf' = 'WhatIfPreference'
        'Debug' = 'DebugPreference'
    }).AsReadOnly()

# Import the XML code for the classic progress window.
$Dialogs.Classic.ProgressWindow.XamlCode = [System.IO.StringReader]::new(@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="Window" Title="" ToolTip="" Padding="0,0,0,0" Margin="0,0,0,0" WindowStartupLocation="Manual" Top="0" Left="0" Topmost="" ResizeMode="NoResize" ShowInTaskbar="True" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" SizeToContent="WidthAndHeight">
    <Window.Resources>
        <Storyboard x:Key="Storyboard1" RepeatBehavior="Forever">
            <DoubleAnimationUsingKeyFrames BeginTime="00:00:00" Storyboard.TargetName="ellipse" Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[2].(RotateTransform.Angle)">
                <SplineDoubleKeyFrame KeyTime="00:00:02" Value="360" />
            </DoubleAnimationUsingKeyFrames>
        </Storyboard>
    </Window.Resources>
    <Window.Triggers>
        <EventTrigger RoutedEvent="FrameworkElement.Loaded">
            <BeginStoryboard Storyboard="{StaticResource Storyboard1}" />
        </EventTrigger>
    </Window.Triggers>
    <Grid Background="#F0F0F0" MinWidth="450" MaxWidth="450" Width="450">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition MinWidth="100" MaxWidth="100" Width="100" />
            <ColumnDefinition MinWidth="350" MaxWidth="350" Width="350" />
        </Grid.ColumnDefinitions>
        <Image x:Name="ProgressBanner" Grid.ColumnSpan="2" Margin="0,0,0,0" Grid.Row="0" />
        <TextBlock x:Name="ProgressText" Grid.Row="1" Grid.Column="1" Margin="0,30,64,30" Text="" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Center" TextAlignment="" Padding="10,0,10,0" TextWrapping="Wrap" />
        <Ellipse x:Name="ellipse" Grid.Row="1" Grid.Column="0" Margin="0,0,0,0" StrokeThickness="5" RenderTransformOrigin="0.5,0.5" Height="32" Width="32" HorizontalAlignment="Center" VerticalAlignment="Center">
            <Ellipse.RenderTransform>
                <TransformGroup>
                    <ScaleTransform />
                    <SkewTransform />
                    <RotateTransform />
                </TransformGroup>
            </Ellipse.RenderTransform>
            <Ellipse.Stroke>
                <LinearGradientBrush EndPoint="0.445,0.997" StartPoint="0.555,0.003">
                    <GradientStop Color="White" Offset="0" />
                    <GradientStop Color="#0078d4" Offset="1" />
                </LinearGradientBrush>
            </Ellipse.Stroke>
        </Ellipse>
    </Grid>
</Window>
'@)

# Send the module's database into the C# code for internal access.
[PSADT.Module.InternalDatabase]::Init($ADT)

# Determine how long the import took.
$ADT.Durations.ModuleImport = [System.DateTime]::Now - $ModuleImportStart
& $Script:CommandTable.'Remove-Variable' -Name ModuleImportStart -Force -Confirm:$false



# SIG # Begin signature block
# MIIuKwYJKoZIhvcNAQcCoIIuHDCCLhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB5OriaUd9lSSbq
# YjTk01BYEBUMFy9/K4V112Bwv+KFUaCCE5UwggWQMIIDeKADAgECAhAFmxtXno4h
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
# ghnsMIIZ6AIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEAr5W7a+ogyFDpjG+46sCPkwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgwFoJW6RBB2h1kT8yBYL49cPoydFGCpdrfaOsHkcC/EMw
# DQYJKoZIhvcNAQEBBQAEggGASRMWgJ1MMK57WDztJbarAvTMvp2ake2ALbx7qeEJ
# oLsn0sk0v0RZ8mZpg9Xxf9wGScf5KONBWSBdy9OsqhLB/7g7WBlt0TDu5D3E60D7
# Bs/AJqkCg1av93NG2TsQzuBiE8S7BPaCJtfIiVDsow3x2Z5CaCHmklL4J2y7M16t
# 2yRXNxm76zBDg6BGcrd53x1LKxhIJi1VYpqWoeF4VFzu+FWN+tD/qCVpZC4s1d0J
# C5h1UDSYkne0JorWsZa/CsvHQpTKbNMbP24YE4+Hsc4ZyFGWbQokyDlooIekzP0x
# ebnHsh2p/ZJdYELnMBwLsUNAygvVrOVQrplb7S2vjkOr610a1cBrUx4sEhaNURcx
# zddMJUs+dSArH5xvQ75mX1tot6X2N+PTH4IoPXZrl/goCElo5GRuXG0U7l6ylQ6K
# q+YyFZxmwFTSEwT1kxiS6QTakMDYzYeqWQb7uM+gP8CmdsK0efrIZ4w+ktllJwqv
# d0LbbVnJu5n4VcKyGY13pzC8oYIXOTCCFzUGCisGAQQBgjcDAwExghclMIIXIQYJ
# KoZIhvcNAQcCoIIXEjCCFw4CAQMxDzANBglghkgBZQMEAgEFADB3BgsqhkiG9w0B
# CRABBKBoBGYwZAIBAQYJYIZIAYb9bAcBMDEwDQYJYIZIAWUDBAIBBQAEIFtMzW/y
# smSZBe7YmDC+ZnEv7G+BZL21DKyNK1jj/Qy6AhBQ2ndrXpxjgON7PRPSuALNGA8y
# MDI0MTIxOTIyNDYwMlqgghMDMIIGvDCCBKSgAwIBAgIQC65mvFq6f5WHxvnpBOMz
# BDANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBT
# SEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTI0MDkyNjAwMDAwMFoXDTM1MTEyNTIz
# NTk1OVowQjELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSAwHgYDVQQD
# ExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyNDCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBAL5qc5/2lSGrljC6W23mWaO16P2RHxjEiDtqmeOlwf0KMCBDEr4I
# xHRGd7+L660x5XltSVhhK64zi9CeC9B6lUdXM0s71EOcRe8+CEJp+3R2O8oo76EO
# 7o5tLuslxdr9Qq82aKcpA9O//X6QE+AcaU/byaCagLD/GLoUb35SfWHh43rOH3bp
# LEx7pZ7avVnpUVmPvkxT8c2a2yC0WMp8hMu60tZR0ChaV76Nhnj37DEYTX9ReNZ8
# hIOYe4jl7/r419CvEYVIrH6sN00yx49boUuumF9i2T8UuKGn9966fR5X6kgXj3o5
# WHhHVO+NBikDO0mlUh902wS/Eeh8F/UFaRp1z5SnROHwSJ+QQRZ1fisD8UTVDSup
# WJNstVkiqLq+ISTdEjJKGjVfIcsgA4l9cbk8Smlzddh4EfvFrpVNnes4c16Jidj5
# XiPVdsn5n10jxmGpxoMc6iPkoaDhi6JjHd5ibfdp5uzIXp4P0wXkgNs+CO/CacBq
# U0R4k+8h6gYldp4FCMgrXdKWfM4N0u25OEAuEa3JyidxW48jwBqIJqImd93NRxvd
# 1aepSeNeREXAu2xUDEW8aqzFQDYmr9ZONuc2MhTMizchNULpUEoA6Vva7b1XCB+1
# rxvbKmLqfY/M/SdV6mwWTyeVy5Z/JkvMFpnQy5wR14GJcv6dQ4aEKOX5AgMBAAGj
# ggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEw
# HwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFJ9XLAN3
# DigVkGalY17uT5IfdqBbMFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3Rh
# bXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRw
# Oi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1l
# U3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggIBAD2tHh92mVvjOIQSR9lD
# kfYR25tOCB3RKE/P09x7gUsmXqt40ouRl3lj+8QioVYq3igpwrPvBmZdrlWBb0Hv
# qT00nFSXgmUrDKNSQqGTdpjHsPy+LaalTW0qVjvUBhcHzBMutB6HzeledbDCzFzU
# y34VarPnvIWrqVogK0qM8gJhh/+qDEAIdO/KkYesLyTVOoJ4eTq7gj9UFAL1UruJ
# KlTnCVaM2UeUUW/8z3fvjxhN6hdT98Vr2FYlCS7Mbb4Hv5swO+aAXxWUm3WpByXt
# gVQxiBlTVYzqfLDbe9PpBKDBfk+rabTFDZXoUke7zPgtd7/fvWTlCs30VAGEsshJ
# mLbJ6ZbQ/xll/HjO9JbNVekBv2Tgem+mLptR7yIrpaidRJXrI+UzB6vAlk/8a1u7
# cIqV0yef4uaZFORNekUgQHTqddmsPCEIYQP7xGxZBIhdmm4bhYsVA6G2WgNFYagL
# DBzpmk9104WQzYuVNsxyoVLObhx3RugaEGru+SojW4dHPoWrUhftNpFC5H7QEY7M
# hKRyrBe7ucykW7eaCuWBsBb4HOKRFVDcrZgdwaSIqMDiCLg4D+TPVgKx2EgEdeoH
# NHT9l3ZDBD+XgbF+23/zBjeCtxz+dL/9NWR6P2eZRi7zcEO1xwcdcqJsyz/JceEN
# c2Sg8h3KeFUCS7tpFk7CrDqkMIIGrjCCBJagAwIBAgIQBzY3tyRUfNhHrP0oZipe
# WzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdp
# Q2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIwMzIzMDAwMDAwWhcNMzcwMzIyMjM1
# OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5
# BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0
# YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAxoY1Bkmz
# wT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCwzIP5WvYRoUQVQl+kiPNo+n3znIkL
# f50fng8zH1ATCyZzlm34V6gCff1DtITaEfFzsbPuK4CEiiIY3+vaPcQXf6sZKz5C
# 3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ7Gnf2ZCHRgB720RBidx8ald68Dd5
# n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7QKxfst5Kfc71ORJn7w6lY2zkpsUd
# zTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/tePc5OsLDnipUjW8LAxE6lXKZYnLvWH
# po9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCYOjgRs/b2nuY7W+yB3iIU2YIqx5K/
# oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9KoRxrOMUp88qqlnNCaJ+2RrOdOqPV
# A+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6dSgkQe1CvwWcZklSUPRR8zZJTYsg
# 0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM1+mYSlg+0wOI/rOP015LdhJRk8mM
# DDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbCdLI/Hgl27KtdRnXiYKNYCQEoAA6E
# VO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbECAwEAAaOCAV0wggFZMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1NhS9zKXaaL3WMaiCPnshvMB8GA1Ud
# IwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNV
# HSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0f
# BDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1
# c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7ZvmKlEIgF+ZtbYIULhsBguEE0TzzBT
# zr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI2AvlXFvXbYf6hCAlNDFnzbYSlm/E
# UExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/tydBTX/6tPiix6q4XNQ1/tYLaqT5Fm
# niye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVPulr3qRCyXen/KFSJ8NWKcXZl2szw
# cqMj+sAngkSumScbqyQeJsG33irr9p6xeZmBo1aGqwpFyd/EjaDnmPv7pp1yr8TH
# wcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc6UsCUqc3fpNTrDsdCEkPlM05et3/
# JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3cHXg65J6t5TRxktcma+Q4c6umAU+9
# Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0dKNPH+ejxmF/7K9h+8kaddSweJywm
# 228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZPJ/tgZxahZrrdVcA6KYawmKAr7ZVB
# tzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLeMt8EifAAzV3C+dAjfwAL5HYCJtnw
# ZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDyDivl1vupL0QVSucTDh3bNzgaoSv2
# 7dZ8/DCCBY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEM
# BQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJ
# RCBSb290IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zC
# pyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf
# 1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x
# 4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEio
# ZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7ax
# xLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZ
# OjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJ
# l2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz
# 2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH
# 4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb
# 5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ
# 9eRpL5gdLfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYD
# VR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuC
# MS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0g
# ADANBgkqhkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs
# 7IVeqRq7IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq
# 3votVs/59PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/
# Lwum6fI0POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9
# /HYJaISfb8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWoj
# ayL/ErhULSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDGCA3YwggNyAgEB
# MHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYD
# VQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFt
# cGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEFAKCB0TAaBgkq
# hkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI0MTIxOTIy
# NDYwMlowKwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU29OF7mLb0j575PZxSFCHJNWG
# W0UwLwYJKoZIhvcNAQkEMSIEID7jb/ahTrcR0Y62eRynk2x1Q2BfWsBSlGryOt0k
# 8y27MDcGCyqGSIb3DQEJEAIvMSgwJjAkMCIEIHZ2n6jyYy8fQws6IzCu1lZ1/tdz
# 2wXWZbkFk5hDj5rbMA0GCSqGSIb3DQEBAQUABIICAEJQwJR0l0EHCvqC6TSFhIiK
# N4oECyPp1LtI1sFUvvN3R/BJdp5D7b41tFaeL9iY8VkmLU/eO9y8jXc73RhT+jEW
# blMTUbzeGPx59vYVd9V790fl82MiXoommojbm3ya/oiRKbSbI7lsYkcfrA8WfO3p
# IwJoU7xCoYISpKHE7DSwWFQKyn7H+l3lCNsA2U4IW50syFu9XvtLRWp56EQx7kKs
# 6KYYZw+OSfe8I3eMZdbCBm/k0Ij4J+p8yDf5KcRyUelLnT6OOGTAYgi9UvTVo0ET
# 13v2t6hCMCh3zHALaSbDprahficrXlY2Kc1vEyZV/M5gZuMQ3/xMktnAzMU0jI3m
# 5wgCo5nf08YNu9/Kg5SnEiMNbIsVr/ZAPkK9hsCk5V08wIhrKVYmJ4yYPienrY5S
# a0EGIv/OOzb3Gg4K3PdiGSgtbK+z8u2Lj4bFMLD88K4vNyt+yJQTVbfAr/BenpuR
# J+g8RMs9Bxl3StEH5cJxaUSajzPksCBUi/c3UWF+UCoWyCCbieAmQpqQQYbDnpgK
# HRep7cI3lgjuApU2uWUPMhO4YvGThKGGTz1Mtc05hEOkOWyKe66IQcvFhKtJMCtQ
# M9NfLkZTwa9P6oJionEGlqSvQmH4Yoet7pFiH9YbRCl0Ut5uu17luPdCXuk73ruG
# UpFVUYrfSkZc2eaal/sB
# SIG # End signature block
