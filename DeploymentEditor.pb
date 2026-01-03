;------------------------------------------------------------------------------------
;- Deployment Editor
;------------------------------------------------------------------------------------
;-- This tool has been developed by TUGI.CH - www.tugi.ch

;------------------------------------------------------------------------------------
;- License
;------------------------------------------------------------------------------------
;-- See LICENSE.txt

;------------------------------------------------------------------------------------
;- Notes
;------------------------------------------------------------------------------------
;-- See TODO.txt

;------------------------------------------------------------------------------------
;- Compiler Settings
;------------------------------------------------------------------------------------
EnableExplicit
UsePNGImageDecoder()
UseSQLiteDatabase()
UseZipPacker()

;------------------------------------------------------------------------------------
;- Structures
;------------------------------------------------------------------------------------
Structure PSADT_Parameter
  ; P.ID, C.Command,P.Parameter,P.Description,P.Type,P.Required
  Command.s
  Parameter.s
  Description.s
  Type.s
  Required.i
EndStructure

Structure PSADT_Parameter_GadgetGroup
  ActionID.i
  Parameter.s
  TitleGadget.i
  InputGadget.i
  ControlType.s
EndStructure

Structure RecentProject
  FileName.s
  FolderPath.s
EndStructure

Structure ProjectSetting
  Value.s
EndStructure

Structure EditorPlugin
  ID.s
  Name.s
  Version.s
  Date.s
  Description.s
  Author.s
  Website.s
  Path.s
  File.s
  Parameter.s
  UnloadProject.s
EndStructure

Structure WinGetPackage
  PackageIdentifier.s
  PackageVersion.s
  Scope.s
  ReleaseDate.s
  File.s
EndStructure

Structure MsiInformation
  Productmanufacturer.s
  Productname.s
  Productcode.s
  Productversion.s
EndStructure

;------------------------------------------------------------------------------------
;- Variables, Enumerations and Maps
;------------------------------------------------------------------------------------
Global Event = #Null, Quit = #False
Global MainWindowTitle.s = "Deployment Editor - TUGI.CH"
Global DonationUrl.s = "https://www.paypal.com/donate/?hosted_button_id=PXABL8ESQQ4F8"
Global PluginDirectory.s = GetCurrentDirectory() + "Plugins"
Global SnippetsDirectory.s = GetCurrentDirectory() + "Snippets"
Global ScriptEditorFilePath.s = GetCurrentDirectory() + "Web\MonacoEditor\editor.html"

; Templates
Global Template_PSADT.s = GetCurrentDirectory() + "ThirdParty\PSAppDeployToolkit\"
Global Template_EmptyDatabase.s = GetCurrentDirectory() + "Templates\Invoke-AppDeployToolkit.db"

; PSADT
Global PSADT_Database.s = GetCurrentDirectory() + "Databases\PSADT.sqlite"
Global PSADT_TemplateFile.s = GetCurrentDirectory() + "Templates\Invoke-AppDeployToolkit.ps1"
Global PSADT_OnlineDocumentation.s = "https://psappdeploytoolkit.com/docs"

; Intune
Global IntuneWinAppUtil.s = ""

; Project
Global Project_FolderPath.s = GetCurrentDirectory() + "Test\"
Global Project_DeploymentFile.s = Project_FolderPath + "Invoke-AppDeployToolkit.ps1"
Global Project_Database.s = Project_FolderPath + "Invoke-AppDeployToolkit.db"
Global Project_PreviewMode = #False

; Windows Sandbox
Global PSADT_SandboxTemplate.s = GetCurrentDirectory() + "Templates\Windows Sandbox.wsb"

; Deployment
Global CurrentDeploymentType.s = "Installation"

; Maps
Global NewMap PSADT_Parameters.PSADT_Parameter()
Global NewMap ProjectSettings.ProjectSetting()

; Dim
Global Dim MsiInformation.MsiInformation(1)

; Lists
Global NewList EditorPlugins.EditorPlugin()
Global NewList WinGetPackages.WinGetPackage()
Global NewList RecentProjects.RecentProject()
AddElement(RecentProjects())
RecentProjects()\FileName = "Invoke-AppDeployToolkit.db"
RecentProjects()\FolderPath = Project_FolderPath

; Action
Global SelectedActionID.i
Global UnsavedChange = #False
Global NewList ActionParameterGadgets.PSADT_Parameter_GadgetGroup()

; WinAPI
Global tvi.TV_ITEM

; WinGet Package
Global Download_Url.s = "", Download_OutputFile.s = ""
Global WinGet_RepositoryUrl.s = "https://github.com/microsoft/winget-pkgs/archive/refs/heads/master.zip"
Global WinGet_RepositoryLocalFile.s = GetTemporaryDirectory() + "WinGetRepository.zip"
Global WinGet_ManifestTempFolder.s = GetTemporaryDirectory() + "Deployment Editor"
Global WinGet_Identifier.s = "", WinGet_Version.s = "", WinGet_SilentSwitch.s = ""

; Script Editor
Global ScriptEditorGadget

;------------------------------------------------------------------------------------
;- Shortcuts
;------------------------------------------------------------------------------------
Enumeration KeyboardShortcuts
  
  ; Menu
  #MenuItem_New
  #MenuItem_Open
  #MenuItem_Save
  #MenuItem_Reload
  #MenuItem_Quit
  #MenuItem_ShowProjectFolder
  #MenuItem_ShowFilesFolder
  #MenuItem_ShowSupportFilesFolder
  #MenuItem_OpenWithISE
  #MenuItem_OpenWithNotepadPlusPlus
  #MenuItem_OpenWithVSCode
  #MenuItem_RunInstallation
  #MenuItem_RunInstallationSandbox
  #MenuItem_RunRemoteMachine
  #MenuItem_RunUninstall
  #MenuItem_RunRepair
  #MenuItem_CreateIntunePackage
  #MenuItem_ShowLogs
  #MenuItem_GenerateExecutablesList
  #MenuItem_ShowPlugins
  #MenuItem_PSADT_OnlineDocumentation
  #MenuItem_PSADT_OnlineVariablesOverview
  #MenuItem_AboutApp
  
  ; Keyboard
  #Keyboard_Shortcut_Save
  #Keyboard_Shortcut_Run
  #Keyboard_Shortcut_Exit
  
  ; WinGet
  #WinGetImport_Enter
  
  ; Window Exit
  #MainWindow_Exit
  
EndEnumeration

; Readme
Global Readme.s = "Deployment Editor simplifies software packaging with PSAppDeployToolkit (www.psappdeploytoolkit.com)." + #CRLF$ + #CRLF$ +
                  "By using this application, you accept the attached license terms of the software. See also LICENSE.txt in the installation folder. Each generated project result is under the license of PSAppDeployToolkit." + #CRLF$ + #CRLF$ +
                  "If you have any questions, please contact us by e-mail: contact@tugi.ch."

;------------------------------------------------------------------------------------
;- Enumerations
;------------------------------------------------------------------------------------
; None

;------------------------------------------------------------------------------------
;- Forms
;------------------------------------------------------------------------------------
XIncludeFile "Forms/MainWindow.pbf"
XIncludeFile "Forms/AboutWindow.pbf"
XIncludeFile "Forms/NewProjectWindow.pbf"
XIncludeFile "Forms/ProjectSettingsWindow.pbf"
XIncludeFile "Forms/PluginWindow.pbf"
XIncludeFile "Forms/WinGetImportWindow.pbf"
XIncludeFile "Forms/ProgressWindow.pbf"
XIncludeFile "Forms/ImportExeWindow.pbf"
XIncludeFile "Forms/ImportMsiWindow.pbf"
XIncludeFile "Forms/ScriptEditorWindow.pbf"

;------------------------------------------------------------------------------------
;- Helpers
;------------------------------------------------------------------------------------

CompilerIf #PB_Compiler_OS = #PB_OS_Windows
  Macro GetParentDirectory(Path)
    GetPathPart(RTrim(Path, "\"))
  EndMacro
CompilerElse
  Macro GetParentDirectory(Path)
    GetPathPart(RTrim(Path, "/"))
  EndMacro
CompilerEndIf

Procedure CheckDatabaseUpdate(Database, Query$)
  Protected Result.i
  Debug "[Debug: Database Update] Update query: " + Query$
  
  Result = DatabaseUpdate(Database, Query$)
  If Result = 0
    Debug DatabaseError()
  Else
    Debug "[Debug: Database Update] Database update was successfully."
    Debug "[Debug: Database Update] Affected rows: " + AffectedDatabaseRows(Database)
  EndIf
  
  ProcedureReturn Result
EndProcedure

;------------------------------------------------------------------------------------
;- Functions
;------------------------------------------------------------------------------------

Procedure DisableMainWindowGadgets(State = #False)
  DisableGadget(ListView_Commands, State)
  DisableGadget(Combo_DeploymentType, State)
  DisableGadget(Tree_Sequence, State)
  DisableGadget(Button_SaveAction, State)
  DisableGadget(Button_AddCommand, State)
  DisableGadget(Button_AddCustomScript, State)
  DisableGadget(String_CommandSearch, State)
  DisableGadget(Hyperlink_ProjectSettings, State)
  DisableGadget(Hyperlink_ProjectSettings, State)
  DisableGadget(ButtonImage_Go, State)
  DisableGadget(Checkbox_SilentMode, State)
  DisableGadget(ButtonImage_StartPowerShell, State)
  DisableGadget(ButtonImage_Repair, State)
  DisableGadget(ButtonImage_Uninstallation, State)
  DisableGadget(ButtonImage_Plugins, State)
  DisableGadget(ButtonImage_RefreshProject, State)
  DisableGadget(ButtonImage_RunHelp, State)
  DisableGadget(ButtonImage_AboutWindow, State)
  DisableGadget(Hyperlink_GenerateDeployment, State)
  DisableGadget(Button_LoadPreview, State)
EndProcedure

Procedure ShowSoftwareReadMe()
  CompilerIf #PB_Compiler_Debugger = 0
    MessageRequester("Readme", Readme, #PB_MessageRequester_Info)
  CompilerEndIf
EndProcedure

Procedure ShowMainWindow()
  Protected ScriptFileName$
  
  OpenMainWindow()
  WindowBounds(MainWindow, WindowWidth(MainWindow)-100, WindowHeight(MainWindow), #PB_Ignore, #PB_Ignore)
  BindEvent(#PB_Event_SizeWindow, @ResizeGadgetsMainWindow(), MainWindow)
  AddKeyboardShortcut(MainWindow, #PB_Shortcut_Escape, #MainWindow_Exit)
  ClearGadgetItems(ListView_Scripts)
  EnableGadgetDrop(Tree_Sequence, #PB_Drop_Files, #PB_Drag_Copy)
  
  ; Fix flickering
  ;SmartWindowRefresh(MainWindow, #False) 
  
  ; Add scripts
  If ExamineDirectory(0, SnippetsDirectory, "*.ps1")

    While NextDirectoryEntry(0)
      ScriptFileName$ = DirectoryEntryName(0)
      If DirectoryEntryType(0) = #PB_DirectoryEntry_File
        AddGadgetItem(ListView_Scripts, -1, ScriptFileName$)
      EndIf
    Wend
    
  Else
    MessageRequester("Error","Can't examine this directory: "+GetGadgetText(0),0)
  EndIf

EndProcedure

Procedure ShowSnippetsFolder(Event)
  RunProgram(SnippetsDirectory)
EndProcedure

Procedure AddSnippet(EventType)
  Protected SelectedScriptFileName.s = GetGadgetText(ListView_Scripts)
  Protected Format, ScriptContent.s
  Protected LastStep.i = 0, NextStep.i = 0, LastID.i = 0
  Protected Command.s = "#CustomScript"
  Protected Query.s, LastResult
  
  
  If SelectedScriptFileName <> ""
    If ReadFile(0, SnippetsDirectory + "\" + SelectedScriptFileName)
      Format = ReadStringFormat(0)
      
      While Eof(0) = 0
        ScriptContent = ScriptContent + ReadString(0, Format) + Chr(13) + Chr(10)
      Wend
      
      CloseFile(0)
    Else
      MessageRequester("Error", "Couldn't open the file: " + SnippetsDirectory + "\" + SelectedScriptFileName, #PB_MessageRequester_Error | #PB_MessageRequester_Ok)
    EndIf
  EndIf
  
  ; Retrieve last step
  SetDatabaseString(1, 0, CurrentDeploymentType)
  If DatabaseQuery(1, "SELECT MAX(Step) FROM Actions WHERE DeploymentType=? LIMIT 1")
    If FirstDatabaseRow(1)
      LastStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Script Custom Script Handler] Last step count: "+LastStep
    EndIf
  EndIf
  
  ; Calc next step
  NextStep = LastStep + 1
  
  ; Add new action
  If IsDatabase(1) And Command <> ""
    
    ; Insert action
    Query.s = "INSERT INTO Actions (Step, Command, Name, Disabled, ContinueOnError, DeploymentType) VALUES ("+NextStep+", '"+Command+"', 'PowerShell script ("+SelectedScriptFileName+")', 0, 0, '"+CurrentDeploymentType+"')"
    Debug "[Debug: Script Custom Script Handler] Update table: "+Query
    CheckDatabaseUpdate(1, Query)
    
    ; Last row
    SetDatabaseString(1, 0, CurrentDeploymentType)
    If DatabaseQuery(1, "SELECT MAX(ID) FROM Actions WHERE DeploymentType=? LIMIT 1")
      If FirstDatabaseRow(1)
        LastID = GetDatabaseLong(1, 0)
        Debug "[Debug: Script Custom Script Handler] Last MAX(ID) is: "+LastID
      EndIf
    EndIf
    
    ; Insert action value
    Query.s = "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES ("+LastID+", 'Script', '"+ScriptContent+"')"
    Debug "[Debug: Script Custom Script Handler] Update table: "+Query
    CheckDatabaseUpdate(1, Query)
  EndIf
  
  FinishDatabaseQuery(1)
  RefreshProject(0)

EndProcedure

Procedure UpdateProjectSettings(SettingName.s, Value.s)
  SetDatabaseString(1, 0, SettingName)
  SetDatabaseString(1, 1, Value)
  CheckDatabaseUpdate(1, "INSERT INTO Settings (Name, Value) VALUES (?, ?)")
EndProcedure

Procedure UpdateProjectSettingByGadget(SettingName.s, Gadget)
  SetDatabaseString(1, 0, GetGadgetText(Gadget))
  SetDatabaseString(1, 1, SettingName)
  CheckDatabaseUpdate(1, "UPDATE Settings SET Value = ? WHERE Name = ?")
EndProcedure

Procedure.s ReplaceDotsAndForwardSlashes(String.s)
  String = ReplaceString(String, Chr(47), Chr(95))
  String = ReplaceString(String, Chr(46), Chr(95))
  ProcedureReturn String
EndProcedure

Procedure DownloadInternetFile(Url.s, OutputFile.s, Gadget, What.s = "file")
  Protected Download, Progress
  
  ; Receive HTTP file from url and save it as local file to the file system
  Download = ReceiveHTTPFile(Url, OutputFile, #PB_HTTP_Asynchronous)
  
  ; Progress download
  If Download
    Debug "[Debug: Internet File Downloader] Starting download ["+Url+"] to: " + OutputFile
    Repeat
      Progress = HTTPProgress(Download)
      Select Progress
          
        ; Success
        Case #PB_HTTP_Success
          SetGadgetText(Gadget, "Download successfully")
          FinishHTTP(Download)
          Break
          
        ; Failed
        Case #PB_HTTP_Failed
          Debug "[Debug: Internet File Downloader] Download failed"
          SetGadgetText(Gadget, "Download failed")
          FinishHTTP(Download)
          DeleteFile(OutputFile)
          ProcedureReturn
          
        ; Aborted
        Case #PB_HTTP_Aborted
          Debug "[Debug: Internet File Downloader] Download aborted"
          SetGadgetText(Gadget, "Download aborted")
          FinishHTTP(Download)
          DeleteFile(OutputFile)
          ProcedureReturn
          
        ; Default
        Default
          Debug "[Debug: Internet File Downloader] Progress of downloading the installer file from internet: " + StrF(Progress / (1024 * 1024), 2) + " MB"
          SetGadgetText(Gadget, "Downloading "+What+" from internet: " + StrF(Progress / (1024 * 1024), 2) + " MB")
          
      EndSelect
      
      Delay(500) ; Don't steal the whole CPU
    ForEver
  EndIf
  
EndProcedure

Procedure.i UpdateLocalWinGetRepository()
  
  If FileSize(WinGet_RepositoryLocalFile) = -1
    Debug "[Debug: WinGet Repository Downloader] Local master copy is not existing"
    
    ; Repo doesnt exist
    If MessageRequester("WinGet Repository", "You need a local copy of the WinGet repository."+Chr(10)+"Do you want To download the latest version automatically from the internet?", #PB_MessageRequester_Warning | #PB_MessageRequester_YesNo) = #PB_MessageRequester_Yes
      Debug "[Debug: WinGet Repository Downloader] Downloading master from the internet"
      DownloadInternetFile(WinGet_RepositoryUrl, WinGet_RepositoryLocalFile, Text_WinGetStatus)
      ProcedureReturn #True
    Else
      Debug "[Debug: WinGet Repository Downloader] Aborted master download"
      ProcedureReturn #False
    EndIf
  Else
    ; Repo exists already
    Protected WinGet_RepositoryFileDate = GetFileDate(WinGet_RepositoryLocalFile, #PB_Date_Modified)
    
    If MessageRequester("WinGet Repository", "You already have a local copy of the WinGet repository ("+FormatDate("%dd.%mm.%yyyy %hh:%ii", WinGet_RepositoryFileDate)+"). Do you want to update or download the latest version from the internet?", #PB_MessageRequester_Info | #PB_MessageRequester_YesNo) = #PB_MessageRequester_Yes
      Debug "[Debug: WinGet Repository Downloader] Updating master on local filesystem"
      DownloadInternetFile(WinGet_RepositoryUrl, WinGet_RepositoryLocalFile, Text_WinGetStatus)
      ProcedureReturn #True
    Else
      Debug "[Debug: WinGet Repository Downloader] Cancelled master update"
      ProcedureReturn #True
    EndIf
  EndIf
  
  ProcedureReturn #False
EndProcedure

Procedure DownloadWinGetManifests()
  Protected PackagePath.s, ManifestFileName.s
  Protected SearchFor.s = LCase(GetGadgetText(String_WinGetPackageSearch))
  Protected ResultsCount.i = 0
  
  ; Delete temp folder
  DeleteDirectory(WinGet_ManifestTempFolder, "*.*", #PB_FileSystem_Recursive)
  
  If 1 = 1
    Debug "[Debug: WinGet Repository Extractor] File received and written to disk. If the remote file was not found, it will contains the webserver error."
      If OpenPack(0, WinGet_RepositoryLocalFile) 
        
        ; List all the entries
        If ExaminePack(0)
          While NextPackEntry(0)
            ;Debug "Name: " + PackEntryName(0) + ", Size: " + PackEntrySize(0)
            
            If FindString(LCase(PackEntryName(0)), SearchFor) And FindString(PackEntryName(0), "installer.yaml")
              ;Debug "Name: " + PackEntryName(0)
              ;Debug "File: " + GetFilePart(PackEntryName(0))
              
              ; Remove unnecessary informations
              PackagePath = ReplaceString(PackEntryName(0), "winget-pkgs-master/manifests/", "")
              PackagePath = Right(PackagePath, Len(PackagePath) - 2)
              
              ; Define file name for manifest
              ManifestFileName = ReplaceDotsAndForwardSlashes(PackagePath)
              
              ; Debug
              ; AddGadgetItem(WinGetImport_ListIcon, 0, PackagePath)
              
              ; Create directory for extraction
              CreateDirectory(WinGet_ManifestTempFolder)
              
              ; Extract manifest
              If UncompressPackFile(0, WinGet_ManifestTempFolder + "\" + ManifestFileName) = -1
                Debug "[Debug: WinGet Repository Extractor] Error: unsuccessful unpacking of file: " + PackEntryName(0)
              EndIf
              
              ; Count and limits
              ResultsCount = ResultsCount + 1
            EndIf
            
            If ResultsCount >= 200
              Debug "[Debug: WinGet Repository Extractor] Limited to 200 extractions"
              Break
            EndIf

          Wend
        EndIf
        
        ClosePack(0)
      EndIf
  Else
    Debug "[Debug: WinGet Repository Download] Failed to download package manifests."
  EndIf
EndProcedure

Procedure ReadWinGetPackages()
  Protected ManifestDirectory.s = WinGet_ManifestTempFolder + "\"
  Protected FileName$, Format, Line.s
  Protected PackageID.s, PackageVersion.s, Scope.s, ReleaseDate.s, File.s
  
  If ExamineDirectory(0, ManifestDirectory, "*.*") 
    While NextDirectoryEntry(0)
      FileName$ = DirectoryEntryName(0)
      
      If DirectoryEntryType(0) = #PB_DirectoryEntry_File
        ;AddGadgetItem(WinGetImport_ListIcon, -1, FileName$)
        
        If ReadFile(0, ManifestDirectory + FileName$)
          Format = ReadStringFormat(0)
          While Eof(0) = 0
            Line = ReadString(0, Format)
            
            If FindString(Line, "PackageIdentifier:") And Not FindString(Line, " - PackageIdentifier:")
              PackageID = Trim(ReplaceString(Line, "PackageIdentifier:", ""))
            EndIf
            
            If FindString(Line, "PackageVersion:")
              PackageVersion = Trim(ReplaceString(Line, "PackageVersion:", ""))
            EndIf
            
            If FindString(Line, "Scope:")
              Scope = Trim(ReplaceString(Line, "Scope:", ""))
            EndIf
            
            If FindString(Line, "ReleaseDate:")
              ReleaseDate = Trim(ReplaceString(Line, "ReleaseDate:", ""))
            EndIf
          Wend
          
          ; Add element to the list
          AddElement(WinGetPackages())
          WinGetPackages()\File = FileName$
          WinGetPackages()\PackageIdentifier = PackageID
          WinGetPackages()\PackageVersion = RemoveString(RemoveString(PackageVersion, Chr(34)), Chr(39))
          WinGetPackages()\Scope = Scope
          WinGetPackages()\ReleaseDate = ReleaseDate
          
          CloseFile(0)
        Else
          MessageRequester("Information", "Couldn't open the file!")
        EndIf
        
        ; Reset values
        PackageID = ""
        PackageVersion = ""
        Scope = ""
        ReleaseDate = ""
      EndIf
    Wend
  Else
    MessageRequester("Error","Can't examine the folder for all manifest files (or 0 search results - Please check the ID).",0)
  EndIf
EndProcedure

Procedure RenderWinGetPackages()
  ; Reset list
  ClearList(WinGetPackages())
  
  ; Read all packages
  ReadWinGetPackages()
  
  ; Clear gadgets
  ClearGadgetItems(WinGetImport_ListIcon)
  
  ; Render in the list icon gadget
  ForEach WinGetPackages()
    AddGadgetItem(WinGetImport_ListIcon, 0, WinGetPackages()\PackageIdentifier +Chr(10)+ 
                                            WinGetPackages()\PackageVersion +Chr(10)+
                                            WinGetPackages()\Scope +Chr(10)+
                                            WinGetPackages()\ReleaseDate +Chr(10)+
                                            WinGetPackages()\File)
  Next
EndProcedure

Procedure SearchWinGetPackage(EventType)
  ; Set status text
  SetGadgetText(Text_WinGetStatus, "Downloading and extracting packages...")
  
  ; Download and extracting manifests
  DownloadWinGetManifests()
  
  ; Render package to the UI
  RenderWinGetPackages()
  
  ; Set final status text
  SetGadgetText(Text_WinGetStatus, "Select the package and press [Create New Project]")
EndProcedure

Procedure CloseWinGetImportWindow(EventType)
  HideWindow(WinGetImportWindow, #True)
EndProcedure

Procedure CloseProgressWindow(EventType)
  HideWindow(ProgressWindow, #True)
  SetGadgetText(Text_ProgressStatus, "No running task.")
EndProcedure

Procedure DownloadInstallerFile(Event)
  Protected Download, Progress
  
  ; Receive HTTP file from url and save it as local file to the file system
  Download = ReceiveHTTPFile(Download_Url, Download_OutputFile, #PB_HTTP_Asynchronous)
  
  ; Progress download
  If Download
    Debug "[Debug: Installer Download] Starting download to: " + Download_OutputFile
    Repeat
      Progress = HTTPProgress(Download)
      Select Progress
          
        ; Success
        Case #PB_HTTP_Success
          SetGadgetText(Text_ProgressStatus, "Download successfully: " +Chr(10)+ 
                                             Download_Url +Chr(10)+Chr(10)+
                                             "Saved file in location: " +Chr(10)+
                                             Download_OutputFile)
          FinishHTTP(Download)
          Break
          
        ; Failed
        Case #PB_HTTP_Failed
          Debug "[Debug: Installer Download] Download failed"
          SetGadgetText(Text_ProgressStatus, "Download failed: " + Download_Url)
          FinishHTTP(Download)
          DeleteFile(Download_OutputFile)
          ProcedureReturn
          
        ; Aborted
        Case #PB_HTTP_Aborted
          Debug "[Debug: Installer Download] Download aborted"
          SetGadgetText(Text_ProgressStatus, "Download aborted: " + Download_Url)
          FinishHTTP(Download)
          DeleteFile(Download_OutputFile)
          ProcedureReturn
          
        ; Default
        Default
          Debug "[Debug: Installer Download] Progress of downloading the installer file from internet: " + StrF(Progress / (1024 * 1024), 2) + " MB"
          SetGadgetText(Text_ProgressStatus, "Downloading the installer file from internet: " + StrF(Progress / (1024 * 1024), 2) + " MB")
          
      EndSelect
      
      Delay(500) ; Don't steal the whole CPU
    ForEver
    
    ; Next steps: Create entry for project
    Delay(1000)
    SetGadgetText(Text_ProgressStatus, "Please wait - Adding installer and silent switch to the project...")
    Delay(1000)
    
    ; Update project database
    If IsDatabase(1)
      If GetExtensionPart(Download_Url) = "msi"
        SetDatabaseString(1, 0, "Start-ADTMsiProcess")
        SetDatabaseString(1, 1, "Start MSI Installer")
        SetDatabaseString(1, 2, "Running the setup installer")
        CheckDatabaseUpdate(1, "INSERT INTO Actions (ID, Step, Command, Name, Disabled, Description, DeploymentType) VALUES (1, 1, ?, ?, 0, ?, 'Installation')")
        CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'FilePath', 'Installer.msi')")
        CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'Action', 'Install')")
        FinishDatabaseQuery(1)
      ElseIf GetExtensionPart(Download_Url) = "exe"
        SetDatabaseString(1, 0, "Start-ADTProcess")
        SetDatabaseString(1, 1, "Start EXE Installer")
        SetDatabaseString(1, 2, "Running the setup installer")
        CheckDatabaseUpdate(1, "INSERT INTO Actions (ID, Step, Command, Name, Disabled, Description, DeploymentType) VALUES (1, 1, ?, ?, 0, ?, 'Installation')")
        CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'FilePath', 'Installer.exe')")
        CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'ArgumentList', '"+WinGet_SilentSwitch+"')")
        FinishDatabaseQuery(1)
      Else
        MessageRequester("WinGet Import", "Sorry, but the downloaded installer is an file type ("+GetExtensionPart(Download_Url)+"), which is Not currently supported For automated sequence generation.", #PB_MessageRequester_Error | #PB_MessageRequester_Ok)
      EndIf
      
      ; Update progress text
      SetGadgetText(Text_ProgressStatus, "The WinGet Import is now complete. You may now continue in the editor.")
      
      ; Refresh the UI
      RefreshProject(0)
      Delay(500)
      UpdateWindow_(MainWindow)
      ;HideWindow(ProgressWindow, #True)
    EndIf
  Else
    Debug "[Debug: Installer Download] Download error: Could not initiate download."
  EndIf
EndProcedure

Procedure PowerShell_ReadMsi(FilePath.s)
  Protected Compiler = #Null
  Protected Output$ = ""
  Protected Exitcode$ = ""
  Protected PSExitcode.i = 1
  Protected Productcode.s = ""

  Compiler = RunProgram("powershell.exe", 
                        "-NoProfile -NoLogo -WindowStyle Hidden -File .\Scripts\Read-Msi.ps1 -FilePath " + Chr(34) + FilePath + Chr(34) + "  -ExecutionPolicy Bypass", 
                        "", 
                        #PB_Program_Open | #PB_Program_Hide | #PB_Program_Read)
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ = ReadProgramString(Compiler)
        
        If FindString(Output$, "Productmanufacturer:")
          MsiInformation(0)\Productmanufacturer = RemoveString(Output$, "Productmanufacturer: ")
        EndIf
        
        If FindString(Output$, "Productname:")
          MsiInformation(0)\Productname = RemoveString(Output$, "Productname: ")
        EndIf
        
        If FindString(Output$, "Productcode:")
          MsiInformation(0)\Productcode = RemoveString(Output$, "Productcode: ")
        EndIf
        
        If FindString(Output$, "Productversion:")
          MsiInformation(0)\Productversion = RemoveString(Output$, "Productversion: ")
        EndIf
      EndIf
    Wend

    PSExitcode = ProgramExitCode(Compiler)
    CloseProgram(Compiler)
  EndIf
  
  If (PSExitcode = 0)
    Debug "[Debug: PowerShell > MSI Reader] " + Productcode
  Else
    MessageRequester("PowerShell Error", Output$, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure LoadProjectSettings()
  
  ; Clear first the map with the old values
  ClearMap(ProjectSettings())
  
  ; Read all new settings from the database
  If OpenDatabase(1, Project_Database, "", "")
    Debug "[Debug: Project Setting] Loaded Project Database successfully: " + PSADT_Database
    
    ; Get all project settings
    If DatabaseQuery(1, "SELECT Name, Value FROM Settings")
      While NextDatabaseRow(1)
        AddMapElement(ProjectSettings(), GetDatabaseString(1, 0))
        ProjectSettings()\Value = GetDatabaseString(1, 1)
      Wend
    Else
      MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
    EndIf
    
    FinishDatabaseQuery(1)
  Else
    MessageRequester("Error", "Can't open the database: " + Project_Database, #PB_MessageRequester_Error)
  EndIf  
EndProcedure

Procedure.s GetProjectSetting(Key.s = "")
  Protected Value.s = ""
  
  If FindMapElement(ProjectSettings(), Key)
    ForEach ProjectSettings()
      If MapKey(ProjectSettings()) = Key
        Value = ProjectSettings()\Value
        Break
      EndIf
    Next
  EndIf
  
  ProcedureReturn Value
EndProcedure

Procedure ShowWinGetImportWindow(EventType)
  If IsWindow(WinGetImportWindow)
    HideWindow(WinGetImportWindow, #False)
    SetActiveWindow(WinGetImportWindow)
  Else
    OpenWinGetImportWindow()
  EndIf
  
  AddKeyboardShortcut(WinGetImportWindow, #PB_Shortcut_Return, #WinGetImport_Enter)
  SetGadgetText(Combo_TargetArchitecture, "x64")
  
  ; Download WinGet repository master
  UpdateLocalWinGetRepository()
EndProcedure

Procedure.s ShortenPathWithDots(Path.s)
  Define ResultPath.s = Path
  Define BackSlashPos.i = 0
  
  ResultPath = Right(ResultPath, 60)
  BackSlashPos = FindString(ResultPath, "\", 0)
  ResultPath = Right(ResultPath, Len(ResultPath) - BackSlashPos)
  
  ProcedureReturn "...\" + RTrim(ResultPath, "\")
EndProcedure

Procedure ShowNewProjectWindow(EventType)    
  If IsWindow(NewProjectWindow)
    HideWindow(NewProjectWindow, #False)
    SetActiveWindow(NewProjectWindow)
  Else
    OpenNewProjectWindow()
  EndIf
  
  ; Hide Gadgets
  ;HideGadget(NPW_Container_ImportWinGet, #True)
  HideGadget(NPW_Container_Recent2, #True)
  HideGadget(NPW_Container_Recent3, #True)
  
  ; Set First Recent Project
  SelectElement(RecentProjects(), 0)
  SetGadgetText(NPW_Recent1_Link_ProjectFile, RecentProjects()\FileName)
  SetGadgetText(NPW_Recent1_Link_ProjectPath, ShortenPathWithDots(RecentProjects()\FolderPath))
  
EndProcedure

Procedure OpenFirstRecentFolder(EventType)
  SelectElement(RecentProjects(), 0)
  RunProgram("explorer.exe", RecentProjects()\FolderPath, "")
EndProcedure

Procedure ProjectIsLoaded()
  If IsDatabase(1)
    ProcedureReturn #True
  EndIf
  
  ; If project is not loaded, show up an error message
  MessageRequester("Error", "No project is currently loaded!" , #PB_MessageRequester_Error | #PB_MessageRequester_Ok)
  ProcedureReturn #False
EndProcedure

Procedure CloseNewProjectWindow(EventType)
  If IsWindow(NewProjectWindow)
    HideWindow(NewProjectWindow, #True)
    HideWindow(MainWindow, #False)
    SetActiveWindow(MainWindow)
  EndIf
EndProcedure

Procedure ShowProjectSettingsWindow(EventType)    
  If IsWindow(ProjectSettingsWindow)
    HideWindow(ProjectSettingsWindow, #False)
    SetActiveWindow(ProjectSettingsWindow)
  Else
    OpenProjectSettingsWindow()
  EndIf
  
  LoadProjectSettings()
  SetGadgetText(PSW_ProjectName_String, GetProjectSetting("Project_Name"))
  SetGadgetText(PSW_AppName_String, GetProjectSetting("App_Name"))
  SetGadgetText(PSW_AppVersion_String, GetProjectSetting("App_Version"))
  SetGadgetText(PSW_AppVendor_String, GetProjectSetting("App_Vendor"))
  SetGadgetText(PSW_AppArch_Combo, GetProjectSetting("App_Architecture"))
  SetGadgetText(PSW_AppLanguage_String, GetProjectSetting("App_Language"))
  SetGadgetText(PSW_ScriptAuthor_String, GetProjectSetting("App_Author"))
EndProcedure

Procedure CloseProjectSettingsWindow(EventType)
  If IsWindow(ProjectSettingsWindow)
    CloseWindow(ProjectSettingsWindow)
    SetActiveWindow(MainWindow)
  EndIf
EndProcedure

Procedure SaveProjectSettings(EventType)
  If IsDatabase(1)
    UpdateProjectSettingByGadget("Project_Name", PSW_ProjectName_String)
    UpdateProjectSettingByGadget("App_Name", PSW_AppName_String)
    UpdateProjectSettingByGadget("App_Version", PSW_AppVersion_String)
    UpdateProjectSettingByGadget("App_Vendor", PSW_AppVendor_String)
    UpdateProjectSettingByGadget("App_Architecture", PSW_AppArch_Combo)
    UpdateProjectSettingByGadget("App_Language", PSW_AppLanguage_String)
    UpdateProjectSettingByGadget("App_Author", PSW_ScriptAuthor_String)
    
    FinishDatabaseQuery(1)
  EndIf
  
  ; Load new values
  LoadProjectSettings()
  
  ; Update Main Window
  SetGadgetText(Text_ProjectName, GetProjectSetting("Project_Name"))
  
  ; Close Window
  CloseProjectSettingsWindow(0)
EndProcedure

Procedure ShowAboutWindow(EventType)
  If IsWindow(AboutWindow)
    HideWindow(AboutWindow, #False)
    SetActiveWindow(AboutWindow)
  Else
    OpenAboutWindow()
  EndIf
EndProcedure

Procedure CloseAboutWindow(EventType)
  If IsWindow(AboutWindow)
    HideWindow(AboutWindow, #True)
    
    Debug "[Debug: Close About Window] Active window ID is: "+GetActiveWindow()
    Debug "[Debug: Close About Window] New project window ID is: "+NewProjectWindow
    
    If GetActiveWindow() = NewProjectWindow
      SetActiveWindow(NewProjectWindow)
    Else
      SetActiveWindow(MainWindow)
    EndIf
  EndIf
EndProcedure

Procedure ShowImportExeWindow(EventType)
  If IsWindow(ImportExeWindow)
    HideWindow(ImportExeWindow, #False)
    SetActiveWindow(ImportExeWindow)
  Else
    OpenImportExeWindow()
  EndIf
EndProcedure

Procedure CloseImportExeWindow(EventType)
  If IsWindow(ImportExeWindow)
    HideWindow(ImportExeWindow, #True)
    
    Debug "[Debug: Close Import Exe Window] Active window ID is: "+GetActiveWindow()
    Debug "[Debug: Close Import Exe Window] New project window ID is: "+NewProjectWindow
    
    If GetActiveWindow() = NewProjectWindow
      SetActiveWindow(NewProjectWindow)
    Else
      SetActiveWindow(MainWindow)
    EndIf
  EndIf
EndProcedure

Procedure ImportExeWindow_SelectFilePath(EventType)
  Protected StandardFile$ = GetCurrentDirectory()
  Protected Pattern$ = "Windows Executable (*.exe)|*.exe|All files (*.*)|*.*"
  Protected Pattern = 0
  Protected ProtectedPattern = 0
  Protected File$ = OpenFileRequester("Please choose the executable to load", StandardFile$, Pattern$, Pattern)
  
  If File$
    SetGadgetText(IEW_String_FilePath, File$)
  EndIf
EndProcedure

Procedure ImportExeWindow_SelectProjectPath(EventType)
  Protected InitialPath$ = "C:\"
  Protected Path$ = PathRequester("Please choose your project path", InitialPath$)
  
  If Path$
    SetGadgetText(IEW_String_ProjectPath, Path$)
  EndIf
EndProcedure

Procedure.i ImportExeWindow_CreateProject(EventType)
  
  ; Set destination path for the new project
  Protected InstallerFile$ = GetGadgetText(IEW_String_FilePath)
  Protected Path$ = GetGadgetText(IEW_String_ProjectPath)
  
  If Path$
    Debug "[Debug: New Project] Choosen path is: " + Path$
  Else
    Debug "[Debug: New Project] Abort project creation - No folder selected."
    ProcedureReturn 0
  EndIf
  
  ; Ask user for confirmation
  Define Confirmation = MessageRequester("Confirmation", "Please confirm the destination folder first - all files will be overwritten with the default template files: " + Path$, #PB_MessageRequester_YesNoCancel | #PB_MessageRequester_Warning)
  If Confirmation = #PB_MessageRequester_No Or Confirmation = #PB_MessageRequester_Cancel
    MessageRequester("Cancelled", "You have canceled the creation of a new project.", #PB_MessageRequester_Ok | #PB_MessageRequester_Info)
    ProcedureReturn #False
  EndIf
  
  ; Copy template folder from PSADT source
  Debug "[Debug: New Project] Copy PSADT framework..."
  Debug "[Debug: New Project] Source Template folder: " + Template_PSADT
  Debug "[Debug: New Project] Destination folder: " + Path$
  CopyDirectory(Template_PSADT, Path$, "", #PB_FileSystem_Recursive | #PB_FileSystem_Force)
  
  ; Copy empty database template to destination folder
  Debug "[Debug: New Project] Copy empty database file..."
  CopyFile(Template_EmptyDatabase, Path$ + "Invoke-AppDeployToolkit.db")
  
  ; Copy installer
  Debug "[Debug: New Project] Copy installer file (" + InstallerFile$ + ") To: " + Path$ + "Files"
  CopyFile(InstallerFile$, Path$ + "Files\" + GetFilePart(InstallerFile$))
  
  ; Set Project Folder and File
  Project_FolderPath.s = Path$
  Project_DeploymentFile.s = Project_FolderPath + "Invoke-AppDeployToolkit.ps1"
  Project_Database.s = Project_FolderPath + "Invoke-AppDeployToolkit.db"
  
  ; Load Project and Settings
  RefreshProject(0)
  LoadProjectSettings()
  
  ; Update project settings
  If IsDatabase(1)
    UpdateProjectSettings("Database_Version", "1.0.4")
    UpdateProjectSettings("Project_Name", "New Import Project")
    UpdateProjectSettings("App_Version", "1.0")
    UpdateProjectSettings("App_Vendor", "Your vendor")
    UpdateProjectSettings("App_Architecture", "x64")
    UpdateProjectSettings("App_Language", "EN")
    UpdateProjectSettings("App_Author", "Executable Import by Deployment Editor")
    UpdateProjectSettings("App_Name", "App name")
    FinishDatabaseQuery(1)
  EndIf
  
  If IsDatabase(1)
    SetDatabaseString(1, 0, "Start-ADTProcess")
    SetDatabaseString(1, 1, "Start Installer")
    SetDatabaseString(1, 2, "Start the installer")
    CheckDatabaseUpdate(1, "INSERT INTO Actions (ID, Step, Command, Name, Disabled, Description, DeploymentType) VALUES (1, 1, ?, ?, 0, ?, 'Installation')")
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'FilePath', '"+GetFilePart(InstallerFile$)+"')")
    FinishDatabaseQuery(1)
    RefreshProject(0)
  EndIf
  
  ; Close New Project Window
  CloseNewProjectWindow(0)
  CloseImportExeWindow(0)
  
  ; Show Main Window
  HideWindow(MainWindow, #False)
  SetActiveWindow(MainWindow)
  SetGadgetText(Text_ProjectName, "New Import Project")
  
  ; Enable GUI
  DisableMainWindowGadgets(#False)
  
  ProcedureReturn #True
EndProcedure

Procedure ShowImportMsiWindow(EventType)
  If IsWindow(ImportMsiWindow)
    HideWindow(ImportMsiWindow, #False)
    SetActiveWindow(ImportMsiWindow)
  Else
    OpenImportMsiWindow()
  EndIf
EndProcedure

Procedure CloseImportMsiWindow(EventType)
  If IsWindow(ImportMsiWindow)
    HideWindow(ImportMsiWindow, #True)
    
    Debug "[Debug: Close Import Msi Window] Active window ID is: "+GetActiveWindow()
    Debug "[Debug: Close Import Msi Window] New project window ID is: "+NewProjectWindow
    
    If GetActiveWindow() = NewProjectWindow
      SetActiveWindow(NewProjectWindow)
    Else
      SetActiveWindow(MainWindow)
    EndIf
  EndIf
EndProcedure

Procedure ImportMsiWindow_SelectFilePath(EventType)
  Protected StandardFile$ = GetCurrentDirectory()
  Protected Pattern$ = "Windows Installer (*.msi)|*.msi|All files (*.*)|*.*"
  Protected Pattern = 0
  Protected ProtectedPattern = 0
  Protected File$ = OpenFileRequester("Please choose the installer to load", StandardFile$, Pattern$, Pattern)
  
  If File$
    SetGadgetText(IMW_String_FilePath, File$)
  EndIf
EndProcedure

Procedure ImportMsiWindow_SelectProjectPath(EventType)
  Protected InitialPath$ = "C:\"
  Protected Path$ = PathRequester("Please choose your project path", InitialPath$)
  
  If Path$
    SetGadgetText(IMW_String_ProjectPath, Path$)
  EndIf
EndProcedure

Procedure.i ImportMsiWindow_CreateProject(EventType)
  
  ; Set destination path for the new project
  Protected InstallerFile$ = GetGadgetText(IMW_String_FilePath)
  Protected Path$ = GetGadgetText(IMW_String_ProjectPath)
  
  If Path$
    Debug "[Debug: New Project] Choosen path is: " + Path$
  Else
    Debug "[Debug: New Project] Abort project creation - No folder selected."
    ProcedureReturn 0
  EndIf
  
  ; Read Msi
  PowerShell_ReadMsi(InstallerFile$)
  
  ; Ask user for confirmation
  Define Confirmation = MessageRequester("Confirmation", "Please confirm the destination folder first - all files will be overwritten with the default template files: " + Path$, #PB_MessageRequester_YesNoCancel | #PB_MessageRequester_Warning)
  If Confirmation = #PB_MessageRequester_No Or Confirmation = #PB_MessageRequester_Cancel
    MessageRequester("Cancelled", "You have canceled the creation of a new project.", #PB_MessageRequester_Ok | #PB_MessageRequester_Info)
    ProcedureReturn #False
  EndIf
  
  ; Copy template folder from PSADT source
  Debug "[Debug: New Project] Copy PSADT framework..."
  Debug "[Debug: New Project] Source Template folder: " + Template_PSADT
  Debug "[Debug: New Project] Destination folder: " + Path$
  CopyDirectory(Template_PSADT, Path$, "", #PB_FileSystem_Recursive | #PB_FileSystem_Force)
  
  ; Copy empty database template to destination folder
  Debug "[Debug: New Project] Copy empty database file..."
  CopyFile(Template_EmptyDatabase, Path$ + "Invoke-AppDeployToolkit.db")
  
  ; Copy installer
  Debug "[Debug: New Project] Copy installer file (" + InstallerFile$ + ") To: " + Path$ + "Files"
  CopyFile(InstallerFile$, Path$ + "Files\" + GetFilePart(InstallerFile$))
  
  ; Set Project Folder and File
  Project_FolderPath.s = Path$
  Project_DeploymentFile.s = Project_FolderPath + "Invoke-AppDeployToolkit.ps1"
  Project_Database.s = Project_FolderPath + "Invoke-AppDeployToolkit.db"
  
  ; Load Project and Settings
  RefreshProject(0)
  LoadProjectSettings()
  
  ; Update project settings
  If IsDatabase(1)
    UpdateProjectSettings("Database_Version", "1.0.4")
    UpdateProjectSettings("Project_Name", "New Import Project")
    UpdateProjectSettings("App_Version", MsiInformation(0)\Productversion)
    UpdateProjectSettings("App_Vendor", MsiInformation(0)\Productmanufacturer)
    UpdateProjectSettings("App_Architecture", "x64")
    UpdateProjectSettings("App_Language", "EN")
    UpdateProjectSettings("App_Author", "MSI Import by Deployment Editor")
    UpdateProjectSettings("App_Name", MsiInformation(0)\Productname)
    FinishDatabaseQuery(1)
  EndIf
  
  If IsDatabase(1)
    SetDatabaseString(1, 0, "Start-ADTMsiProcess")
    SetDatabaseString(1, 1, "Run MSI installer")
    SetDatabaseString(1, 2, "Start the installer")
    CheckDatabaseUpdate(1, "INSERT INTO Actions (ID, Step, Command, Name, Disabled, Description, DeploymentType) VALUES (1, 1, ?, ?, 0, ?, 'Installation')")
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'FilePath', '"+GetFilePart(InstallerFile$)+"')")
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (1, 'Action', 'Install')")
    FinishDatabaseQuery(1)
    
    SetDatabaseString(1, 0, "Start-ADTMsiProcess")
    SetDatabaseString(1, 1, "Run MSI uninstall")
    SetDatabaseString(1, 2, "Start the uninstall")
    CheckDatabaseUpdate(1, "INSERT INTO Actions (ID, Step, Command, Name, Disabled, Description, DeploymentType) VALUES (2, 1, ?, ?, 0, ?, 'Uninstall')")
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (2, 'ProductCode', '"+MsiInformation(0)\Productcode+"')")
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (2, 'Action', 'Uninstall')")
    FinishDatabaseQuery(1)
    
    RefreshProject(0)
  EndIf
  
  ; Close New Project Window
  CloseNewProjectWindow(0)
  CloseImportMsiWindow(0)
  
  ; Show Main Window
  HideWindow(MainWindow, #False)
  SetActiveWindow(MainWindow)
  SetGadgetText(Text_ProjectName, "New Import Project")
  
  ; Enable GUI
  DisableMainWindowGadgets(#False)
  
  ProcedureReturn #True
EndProcedure

Procedure ShowLicensing(EventType)
  RunProgram("notepad.exe", GetCurrentDirectory() + "LICENSE", GetCurrentDirectory())
EndProcedure

Procedure OpenDonationUrl(EventType)
  RunProgram(DonationUrl, "", "")
EndProcedure

Procedure UpdateScript(JsonParameters$)
  Protected EditorContent$ = Mid(JsonParameters$, 3, Len(JsonParameters$) - 4)
  
  ;EditorContent$ = ReplaceString(EditorContent$, "\r", Chr(13))
  ;EditorContent$ = ReplaceString(EditorContent$, "\n", Chr(10))
  ;EditorContent$ = ReplaceString(EditorContent$, Chr(92) + Chr(34), Chr(34))
  
  SetGadgetText(ScriptEditorGadget, UnescapeString(EditorContent$))
  ;ProcedureReturn UTF8(~"150")
EndProcedure

Procedure CloseScriptEditorWindow(EventType)
  If IsWindow(ScriptEditorWindow)
    HideWindow(ScriptEditorWindow, #True)
    HideWindow(MainWindow, #False)
    SetActiveWindow(MainWindow)
  EndIf
EndProcedure

Procedure LoadScriptInEditor()
  Debug "[Debug: Script Editor] Loading Monaco Editor..."
  
  Protected Script$ = GetGadgetText(ScriptEditorGadget)
  HideWindow(ScriptEditorWindow, #False)
  
  WebViewExecuteScript(WebView_ScriptEditor, "myMonacoEditor.setValue('"+EscapeString(Script$)+"')")
  WebViewExecuteScript(WebView_ScriptEditor, "myMonacoEditor.updateOptions({ readOnly: false });")
  BindWebViewCallback(WebView_ScriptEditor, "updateScript", @UpdateScript())
EndProcedure

Procedure EmptyCallback(JsonParameters$)
  Debug "[Empty Callback executed]"
EndProcedure

Procedure NotAvailableFeatureMessage(EventType)
  MessageRequester("Feature not available", "This feature is not yet available for this version. Sign up for the newsletter on the website to be kept up to date.", #PB_MessageRequester_Info | #PB_MessageRequester_Ok) 
EndProcedure

Procedure ShowSoftwareLogFolder(EventType)
  Debug "[Debug: Software Log] Start Windows Explorer for Software Logs."
  RunProgram("explorer.exe", "C:\Windows\Logs\Software", "")
EndProcedure

Procedure ShowProjectFolder(EventType)
  RunProgram("explorer.exe", Project_FolderPath, "")
EndProcedure

Procedure ShowFilesFolder(EventType)
  RunProgram("explorer.exe", Project_FolderPath + "Files\", "")
EndProcedure

Procedure ShowSupportFilesFolder(EventType)
  RunProgram("explorer.exe", Project_FolderPath + "SupportFiles\", "")
EndProcedure

Procedure ShowOnlineDocumentation(EventType)
  RunProgram(PSADT_OnlineDocumentation + "/reference", "", "")
EndProcedure

Procedure ShowOnlineVariablesOverview(EventType)
  RunProgram(PSADT_OnlineDocumentation + "/reference/variables", "", "")
EndProcedure

Procedure ShowCommandHelp(EventType)
  Protected CurrentCommand.s = GetGadgetText(String_ActionCommand)
  RunProgram(PSADT_OnlineDocumentation + "/reference/functions/"+CurrentCommand, "", "")
EndProcedure

Procedure FilterCommands(Search.s)
  Protected PSADT_Command.s
  ClearGadgetItems(ListView_Commands)
  
  ; Filter in SQL query
  If DatabaseQuery(0, "SELECT Command,Category FROM Commands WHERE Command like '%"+Search+"%' ORDER BY Name ASC")
    While NextDatabaseRow(0)
      PSADT_Command = GetDatabaseString(0, 0)
      
      If FindString(PSADT_Command, "#", 1) = 0
        AddGadgetItem(ListView_Commands, -1, PSADT_Command)
      EndIf
    Wend
  Else
    MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure TreeSequence_SetFirstLevelBold()
  Protected a.i
  
  For a = 0 To CountGadgetItems(Tree_Sequence) - 1
    tvi\mask = #TVIF_HANDLE | #TVIF_CHILDREN
    tvi\hItem = GadgetItemID(Tree_Sequence, a)
    SendMessage_(GadgetID(Tree_Sequence), #TVM_GETITEM, 0, @tvi) 
    If tvi\cChildren
      tvi\mask = #TVIF_HANDLE | #TVIF_STATE 
      tvi\state = #TVIS_BOLD 
      tvi\stateMask = #TVIS_BOLD 
      SendMessage_(GadgetID(Tree_Sequence), #TVM_SETITEM, 0, tvi) 
    EndIf
  Next
EndProcedure

Procedure TreeSequence_AllExpanded(State = #PB_Tree_Expanded)
  Protected i.i
  
  For i = 0 To CountGadgetItems(Tree_Sequence) - 1
    SetGadgetItemState(Tree_Sequence, i, State)
  Next i
EndProcedure

Procedure LoadProjectFile()
  Protected CurrentStep.i = 0, LastStep.i = -1
  
  ; Close database if already opened
  If IsDatabase(1)
    CloseDatabase(1)
  EndIf
  
  ; Open the database
  If OpenDatabase(1, Project_Database, "", "")
    Debug "[Debug: Project File Loader] Loaded Project Database successfully: " + Project_Database
    
    SetDatabaseString(1, 0, CurrentDeploymentType)
    If DatabaseQuery(1, "SELECT ID,Step,Command,Name,Disabled,Description,ContinueOnError,Action,Parameter,Value,DeploymentType,VariableName FROM View_Sequence WHERE DeploymentType=? ORDER BY Step ASC")
      
      ; Add steps into the treeview
      While NextDatabaseRow(1)
        
        ; Formatting
        Protected Prefix.s = ""
        Protected IsDisabled = #False
        Protected IconImage = ImageID(Img_MainWindow_7)
        Protected Affix.s = ""
        
        CurrentStep = GetDatabaseLong(1, 1)
        ;Debug "Current step is (" + CurrentStep + ") And last step was (" + LastStep + ")"

        If CurrentStep <> LastStep
          ; Styles by status
          If GetDatabaseLong(1, 4) = 1
            Prefix = "Disabled | " 
            IsDisabled = #True
            IconImage = ImageID(Img_MainWindow_8)
          EndIf
          
          ; PowerShell custom script?
          If GetDatabaseString(1, 2) = "#CustomScript"
            IconImage = ImageID(Img_MainWindow_12)
          EndIf
          
          ; Contine on Error?
          Protected ContinueOnErrorEnabled = GetDatabaseLong(1, 6)
          If ContinueOnErrorEnabled
            Affix + " (Continue on Error)"
          EndIf
          
          ; Title
          AddGadgetItem(Tree_Sequence, -1, Space(1) + Prefix + GetDatabaseString(1, 3) + Affix + Space(1), IconImage)
          SetGadgetItemData(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, GetDatabaseLong(1, 0))
          
          ; Coloring
          If IsDisabled
            SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_FrontColor, RGB(168, 168, 168))
            SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_BackColor, RGB(245, 245, 245))
          Else
            If GetDatabaseString(1, 2) = "#CustomScript"
              SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_FrontColor, RGB(6, 21, 17))
              SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_BackColor, RGB(202, 231, 247))
            Else
              SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_FrontColor, RGB(6, 21, 17))
              SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_BackColor, RGB(218, 242, 236))
            EndIf
          EndIf
          
          ; Font
          If LoadFont(999, "Segoe UI", 11)
            SetGadgetFont(Tree_Sequence, FontID(999))   ; Set the loaded Courier 10 font as new standard
          EndIf
          
          ; Description as first item
          AddGadgetItem(Tree_Sequence, -1, "Description: " + GetDatabaseString(1, 5), 0, 1)
          
          ; Variable as second item
          If Trim(GetDatabaseString(1, 11)) <> ""
            AddGadgetItem(Tree_Sequence, -1, "Variable: " + GetDatabaseString(1, 11), 0, 1)
          EndIf
          
        EndIf
        
        ; Check value
        Protected Value.s = GetDatabaseString(1, 9)
        If RemoveString(Value, " ") = ""
          Value = "(none)"
        EndIf
        
        If GetDatabaseString(1, 8) <> ""
          AddGadgetItem(Tree_Sequence, -1, GetDatabaseString(1, 8) + ": " + Value, 0, 1)
          
          If Value = "(none)"
            SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_FrontColor, RGB(99, 99, 107))
          EndIf
        Else
          AddGadgetItem(Tree_Sequence, -1, "No values for parameters found", 0, 1)
          SetGadgetItemColor(Tree_Sequence, CountGadgetItems(Tree_Sequence) - 1, #PB_Gadget_BackColor, RGB(255, 255, 223))
        EndIf
        
        LastStep = CurrentStep
        Prefix = ""
      Wend
      
      FinishDatabaseQuery(1) 
    Else
      MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
    EndIf
  Else
    MessageRequester("Database Error", "Can't open the database: " + Project_Database, #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure TreeSequence_ExpandedByID(ID.i = 0)
  Protected Index.i = 0
  
  For Index = 0 To CountGadgetItems(Tree_Sequence)
    Protected Item.i = GetGadgetItemData(Tree_Sequence, Index)
    
    If Item = ID
      ProcedureReturn SetGadgetItemState(Tree_Sequence, Index, #PB_Tree_Expanded)
    EndIf
  Next
EndProcedure

Procedure TreeSequence_DropHandler(Files.s)
  Debug "[Debug: Sequence View Drop Handler] " + Files
  
  Protected LastStep.i = 0, NextStep.i = 0, LastID.i = 0
  Protected ActionName.s = "New Action by File Drop", ActionCommand.s = "Start-ADTProcess", CommandParameter.s = "FilePath"
  
  ; Remove project path in file path
  Files = ReplaceString(Files, Project_FolderPath, "")
  
  ; Retrieve last step
  SetDatabaseString(1, 0, CurrentDeploymentType)
  If DatabaseQuery(1, "SELECT MAX(Step) FROM Actions WHERE DeploymentType=? LIMIT 1")
    If FirstDatabaseRow(1)
      LastStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Add Action Custom Script Handler] Last step count: "+LastStep
    EndIf
  EndIf
  
  ; Calc next step
  NextStep = LastStep + 1
  
  ; Add new action
  If IsDatabase(1)
    
    ; Check extension
    Select GetExtensionPart(Files)
      Case "lnk"
        ActionName = "Remove Shortcut ("+Files+")"
        ActionCommand = "Remove-ADTFile"
        CommandParameter = "Path"
    EndSelect
    
    ; Update database
    CheckDatabaseUpdate(1, "INSERT INTO Actions (Step, Command, Name, Disabled, ContinueOnError, DeploymentType) VALUES ("+NextStep+", '"+ActionCommand+"', '"+ActionName+"', 0, 0, '"+CurrentDeploymentType+"')")
    
    If DatabaseQuery(1, "SELECT MAX(ID) FROM Actions WHERE DeploymentType='"+CurrentDeploymentType+"' LIMIT 1")
      If FirstDatabaseRow(1)
        LastID = GetDatabaseLong(1, 0)
        Debug "[Debug: Last Action] Last id is: "+LastID
      EndIf
    EndIf
    
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES ("+LastID+", '"+CommandParameter+"', '"+Files+"')")
  EndIf
  
  FinishDatabaseQuery(1)
  RefreshProject(0)
EndProcedure

Procedure RefreshProject(EventType)
  ClearGadgetItems(Tree_Sequence)
  LoadProjectFile()
  TreeSequence_SetFirstLevelBold()
  ;TreeSequence_AllExpanded()
EndProcedure

Procedure LoadCommandsAndParameters()
  Protected PSADT_Command.s, PSADT_Category.s
  
  If OpenDatabase(0, PSADT_Database, "", "")
    Debug "[Debug: PSADT Database Loader] Loaded PSADT Database successfully: " + PSADT_Database
    
    ; Get all commands
    If DatabaseQuery(0, "SELECT Command,Category FROM Commands ORDER BY Name ASC")
      While NextDatabaseRow(0)
        PSADT_Command = GetDatabaseString(0, 0)
        PSADT_Category = GetDatabaseString(0, 1)
        
        If FindString(PSADT_Command, "#", 1) = 0
          AddGadgetItem(ListView_Commands, -1, PSADT_Command)
        EndIf
      Wend
    Else
      MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
    EndIf
    
    ; Get all parameters
    If DatabaseQuery(0, "SELECT P.ID, C.Command,P.Parameter,P.Description,P.Type,P.Required FROM Parameters AS P LEFT JOIN Commands AS C ON P.Command = C.ID")
      While NextDatabaseRow(0)
        Define Parameter_ID = GetDatabaseLong(0, 0)
        
        PSADT_Parameters(Str(Parameter_ID))\Command = GetDatabaseString(0, 1)
        PSADT_Parameters()\Parameter = GetDatabaseString(0, 2)
        PSADT_Parameters()\Description = GetDatabaseString(0, 3)
        PSADT_Parameters()\Type = GetDatabaseString(0, 4)
        PSADT_Parameters()\Required = GetDatabaseLong(0, 5)
      Wend
    Else
      MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
    EndIf
    
    FinishDatabaseQuery(0)
  Else
    MessageRequester("Error", "Can't open the database: " + PSADT_Database, #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure.s ParameterTypeByCommand(CommandToSearch.s = "", ParameterToSearch.s = "")
  ;Protected NewMap Result.PSADT_Parameter()
  Protected MyParameterType.s = ""
  
  ForEach PSADT_Parameters()
    ;Debug PSADT_Parameters()\Command + " : " + PSADT_Parameters()\Type
    If PSADT_Parameters()\Command = CommandToSearch And PSADT_Parameters()\Parameter = ParameterToSearch
      Protected RKey.s = MapKey(PSADT_Parameters())
      MyParameterType = PSADT_Parameters()\Type
      Break
      ;Debug "Found parameter for " + Command + ": " + PSADT_Parameters()\Parameter
      ;Result(RKey)\Command = PSADT_Parameters()\Command
      ;Result(RKey)\Parameter = PSADT_Parameters()\Parameter
      ;Result(RKey)\Description = PSADT_Parameters()\Description
      ;Result(RKey)\Type = PSADT_Parameters()\Type
      ;Result(RKey)\Required = PSADT_Parameters()\Required
    EndIf
  Next
  
  ProcedureReturn MyParameterType
EndProcedure

Procedure LoadUI()
  
  ; Output in debug console
  Debug "[Debug: UI] Triggered LoadUI()"
  
  ; Update gadgets and hide image placeholders
  SetGadgetState(Combo_DeploymentType, 0)
  HideGadget(Image_Check, #True)
  HideGadget(Image_DisabledStep, #True)
  HideGadget(Image_PowerShell, #True)
  HideGadget(Image_Exclamation, #True)
  
  ; Load available PowerShell commands and the project
  LoadCommandsAndParameters()
  ;Debug ParameterTypeByCommand("Remove-File", "Path")
  
  ; Set keyboard shortcuts for Main Window
  AddKeyboardShortcut(MainWindow, #PB_Shortcut_Control | #PB_Shortcut_S, #Keyboard_Shortcut_Save)
  AddKeyboardShortcut(MainWindow, #PB_Shortcut_F5, #Keyboard_Shortcut_Run)
  
  ; Set version in window title for Main Window
  SetWindowTitle(MainWindow, MainWindowTitle)
  
EndProcedure

Procedure ListFilesRecursive(Dir.s, List Files.s())
  Protected D
  NewList Directories.s()
  
  If Right(Dir, 1) <> "\"
    Dir + "\"
  EndIf
  
  D = ExamineDirectory(#PB_Any, Dir, "")
  While NextDirectoryEntry(D)
    
    Select DirectoryEntryType(D)
      Case #PB_DirectoryEntry_File
        AddElement(Files())
        Files() = Dir + DirectoryEntryName(D)
      Case #PB_DirectoryEntry_Directory
        Select DirectoryEntryName(D)
          Case ".", ".."
            Continue
          Default
            AddElement(Directories())
            Directories() = Dir + DirectoryEntryName(D)
        EndSelect
    EndSelect
  Wend
  
  FinishDirectory(D)
  
  ForEach Directories()
    ListFilesRecursive(Directories(), Files())
  Next
  
EndProcedure

Procedure GenerateExecutablesList(EventType)
  
  ; Check if the project is loaded
  If Not ProjectIsLoaded() : ProcedureReturn 0 : EndIf
  
  ; Set variables and paths
  Protected InstallationPath.s = PathRequester("Path to the installation folder", "C:\Program Files")
  Protected ExecutablesList.s = ""
  
  If InstallationPath
    NewList F.s()
    ListFilesRecursive(InstallationPath, F())
    
    ForEach F()
      If GetExtensionPart(F()) = "exe"
        ExecutablesList = ExecutablesList + "," + GetFilePart(F(), #PB_FileSystem_NoExtension)
      EndIf
    Next
    
    ExecutablesList = LTrim(ExecutablesList, ",")
    InputRequester("Executables List", "Result:", ExecutablesList)
  EndIf

EndProcedure

Procedure SaveAndCloseProject()
  If IsDatabase(1)
    SaveAction(0)
    CloseDatabase(1)
  EndIf
EndProcedure

Procedure CreateIntunePackage(EventType)
  
  ; First close the project database
  SaveAndCloseProject()
  Delay(300)
  
  ; Set variables and paths
  Protected InstallerFile.s, InstallerPath.s, PackagePath.s
  Protected WrapperParameters$, Compiler, Output$, Exitcode = 1

  InstallerFile = "Invoke-AppDeployToolkit.exe"
  InstallerPath = Project_FolderPath + InstallerFile
  PackagePath = PathRequester("Intune Package Destination", "C:\")
  
  ; Check package path if empty
  If Trim(PackagePath) = ""
    RefreshProject(0)
    ProcedureReturn 0
  EndIf
  
  ; Check if IntuneWinAppUtil.exe exists in the user temp folder
  If FileSize(IntuneWinAppUtil) = -1
    StatusBarText(0, 0, "Error: IntuneWinAppUtil missing")
    RefreshProject(0)
    ProcedureReturn MessageRequester("Win32 Content Prep Tool", IntuneWinAppUtil + " is missing!", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
  
  ; Check if *.intunewin file with same name exists in the package folder
  If ExamineDirectory(0, PackagePath, "*.intunewin")
    While NextDirectoryEntry(0)
      If DirectoryEntryType(0) = #PB_DirectoryEntry_File        
        Protected SearchedPackageFile.s = GetFilePart(InstallerFile, #PB_FileSystem_NoExtension) + ".intunewin"

        If DirectoryEntryName(0) = SearchedPackageFile
          RefreshProject(0)
          ProcedureReturn MessageRequester("Existing Intune package file", "Please delete first the existing package to continue: " + DirectoryEntryName(0), #PB_MessageRequester_Ok | #PB_MessageRequester_Warning)
        EndIf
      EndIf
    Wend
    FinishDirectory(0)
  EndIf

  ; Define parameters for IntuneWinAppUtil
  WrapperParameters$ = "-c %FolderPath% -s %InstallerPath% -o %PackagePath%"
  WrapperParameters$ = ReplaceString(WrapperParameters$, "%FolderPath%", Chr(34) + RTrim(Project_FolderPath, "\") + Chr(34))
  WrapperParameters$ = ReplaceString(WrapperParameters$, "%InstallerPath%", Chr(34) + InstallerPath + Chr(34))
  WrapperParameters$ = ReplaceString(WrapperParameters$, "%PackagePath%", Chr(34) + PackagePath + Chr(34))
  
  Debug "[Debug: Create Intune Package] Parameters for the wrapper: " + WrapperParameters$
  
  ; Run IntuneWinAppUtil
  StatusBarText(0, 0, "Create Intune package...")
  Compiler = RunProgram(IntuneWinAppUtil, WrapperParameters$, "", #PB_Program_Open) ; #PB_Program_Open | #PB_Program_Read | #PB_Program_Hide
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ + ReadProgramString(Compiler) + Chr(13)
      EndIf
    Wend
    
    Exitcode = ProgramExitCode(Compiler)
    Output$ + Chr(13) + Chr(13)
    Output$ + "Exitcode: " + Str(Exitcode)
    
    CloseProgram(Compiler)
    StatusBarText(0, 0, "Done.")
  EndIf

  ; Check the exit code
  If Exitcode = 0
    Debug "[Debug: Create Intune Package] Successfully created the Intune package."
    MessageRequester("Intune Package Creation", "The Intune package has been successfully created: "+PackagePath, #PB_MessageRequester_Info | #PB_MessageRequester_Ok)
  Else
    Debug "[Debug: Create Intune Package] Error Intune package creation!"
    MessageRequester("Intune Package Creation", "The Intune package compiler failed.", #PB_MessageRequester_Error | #PB_MessageRequester_Ok)
  EndIf
  
  ; Load project again
  RefreshProject(0)
  
EndProcedure

Procedure OpenFirstRecentProject(EventType)
  
  ; Set Project Folder and File
  SelectElement(RecentProjects(), 0)
  Project_FolderPath.s = RecentProjects()\FolderPath
  Project_DeploymentFile.s = Project_FolderPath + "Invoke-AppDeployToolkit.ps1"
  Project_Database.s = Project_FolderPath + RecentProjects()\FileName
  
  ; Load Project and Settings
  RefreshProject(0)
  LoadProjectSettings()
  
  ; Close New Project Window
  CloseNewProjectWindow(0)
  
  ; Show Main Window
  HideWindow(MainWindow, #False)
  SetActiveWindow(MainWindow)
  SetGadgetText(Text_ProjectName, GetProjectSetting("Project_Name"))
  
  ; Enable GUI
  DisableMainWindowGadgets(#False)
  
EndProcedure

Procedure OpenOtherProject(EventType)
  
  ; Request file location
  Protected StandardFile$, Pattern$, File$, Pattern
  StandardFile$ = "C:\"
  Pattern$ = "Project Database (*.db)|*.db|All files (*.*)|*.*"
  Pattern = 0
  File$ = OpenFileRequester("Please choose project database file to load", StandardFile$, Pattern$, Pattern)
  
  If File$ = ""
    Debug "[Debug: Project File Selector] Canceled project database selection."
    ProcedureReturn 0
  EndIf
  
  ; Set Project Folder and File
  Project_FolderPath.s = GetPathPart(File$)
  Project_DeploymentFile.s = Project_FolderPath + "Invoke-AppDeployToolkit.ps1"
  Project_Database.s = Project_FolderPath + GetFilePart(File$)
  
  ; Load Project and Settings
  RefreshProject(0)
  LoadProjectSettings()
  
  ; Close New Project Window
  CloseNewProjectWindow(0)
  
  ; Show Main Window
  HideWindow(MainWindow, #False)
  SetActiveWindow(MainWindow)
  SetGadgetText(Text_ProjectName, GetProjectSetting("Project_Name"))
  
  ; Enable GUI
  DisableMainWindowGadgets(#False)
  
EndProcedure

Procedure.i CreateNewProject(EventType)
  
  ; Set destination path for the new project
  Protected InitialPath$, Path$, ExamineFolder
  InitialPath$ = "C:\"
  Path$ = PathRequester("Please choose your path for the new project", InitialPath$)
  
  If Path$
    Debug "[Debug: New Project] Choosen path is: " + Path$
  Else
    Debug "[Debug: New Project] Abort project creation - No folder selected."
    ProcedureReturn 0
  EndIf
  
  ; Ask user for confirmation
  Define Confirmation = MessageRequester("Confirmation", "Please confirm the destination folder first - all files will be overwritten with the default template files: " + Path$, #PB_MessageRequester_YesNoCancel | #PB_MessageRequester_Warning)
  If Confirmation = #PB_MessageRequester_No Or Confirmation = #PB_MessageRequester_Cancel
    MessageRequester("Cancelled", "You have canceled the creation of a new project.", #PB_MessageRequester_Ok | #PB_MessageRequester_Info)
    ProcedureReturn #False
  EndIf
  
  ; Copy template folder from PSADT source
  Debug "[Debug: New Project] Copy PSADT framework..."
  Debug "[Debug: New Project] Source Template folder: " + Template_PSADT
  Debug "[Debug: New Project] Destination folder: " + Path$
  CopyDirectory(Template_PSADT, Path$, "", #PB_FileSystem_Recursive | #PB_FileSystem_Force)
  
  ; Copy empty database template to destination folder
  Debug "[Debug: New Project] Copy empty database file..."
  CopyFile(Template_EmptyDatabase, Path$ + "Invoke-AppDeployToolkit.db")
  
  ; Set Project Folder and File
  Project_FolderPath.s = Path$
  Project_DeploymentFile.s = Project_FolderPath + "Invoke-AppDeployToolkit.ps1"
  Project_Database.s = Project_FolderPath + "Invoke-AppDeployToolkit.db"
  
  ; Load Project and Settings
  RefreshProject(0)
  LoadProjectSettings()
  
  ; Close New Project Window
  CloseNewProjectWindow(0)
  
  ; Show Main Window
  HideWindow(MainWindow, #False)
  SetActiveWindow(MainWindow)
  SetGadgetText(Text_ProjectName, GetProjectSetting("Project_Name"))
  
  ; Enable GUI
  DisableMainWindowGadgets(#False)
  
  ProcedureReturn #True
EndProcedure

Procedure.s GetActionParameterValueByID(ActionID.i, Parameter.s)
  Protected Value.s = ""
  Protected RequestQuery.s = "SELECT Value FROM Actions_Values WHERE Parameter='"+Parameter+"' AND Action="+ActionID
  
  Debug "[Debug: Action Query] Query is: " + RequestQuery
  
  If DatabaseQuery(1, RequestQuery)
    While NextDatabaseRow(1)
      Value = GetDatabaseString(1, 0)
    Wend
    
    FinishDatabaseQuery(1) 
  Else
    MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
  EndIf
  
  ProcedureReturn Value
EndProcedure

Procedure EditorInputHandler()
  Debug "[Debug: Editor Input Handler] Input in Editor > Action triggered."
  UnsavedChange = #True
EndProcedure

Procedure OptionsInputHandler()
  Debug "[Debug: Options Input Handler] Input in Editor > Options triggered."
  SaveAction(0)
EndProcedure

Procedure RenderActionEditor(ID.i = -1)
  
  ; Default values
  Protected SelectedItem.i = GetGadgetState(Tree_Sequence)
  Protected Command.s = "", Name.s = "", Description.s = "", VariableName.s = "", Conditions.s = ""
  Protected Disabled.i = 0, ContinueOnError.i = 0
  Protected ActionScrollAreaGadget = 0
  
  If ID = -1
    ID = GetGadgetItemData(Tree_Sequence, SelectedItem)
  EndIf
  
  ; Check for previous changes
  Protected UnsavedChangeMessage
  If UnsavedChange
    UnsavedChange = #False
    UnsavedChangeMessage = MessageRequester("Save Changes", "Do you want to save your previous change?", #PB_MessageRequester_Warning | #PB_MessageRequester_YesNoCancel)
    
    If UnsavedChangeMessage = #PB_MessageRequester_Yes
      SaveAction(0)
      TreeSequence_ExpandedByID(ID)
    EndIf
  EndIf

  Debug "[Debug: Action Editor Renderer] You clicked on item with data ID: " + ID
  SelectedActionID = ID
  
  If ID = 0
    Debug "[Debug: Action Editor Renderer] No action rendering by ID=0."
    SetGadgetText(String_ActionCommand, "")
    SetGadgetText(String_ActionName, "")
    SetGadgetText(String_ActionVariable, "")
    SetGadgetText(Editor_ActionDescription, "")
    SetGadgetText(Editor_ActionConditions, "")
    SetGadgetState(Checkbox_ContinueOnError, #PB_Checkbox_Unchecked)
    SetGadgetState(Checkbox_DisabledAction, #PB_Checkbox_Unchecked)
    
    ; Disable gadgets
    DisableGadget(String_ActionName, #True)
    DisableGadget(Editor_ActionDescription, #True)
    DisableGadget(Editor_ActionConditions, #True)
    DisableGadget(String_ActionVariable, #True)
    DisableGadget(Checkbox_ContinueOnError, #True)
    DisableGadget(Checkbox_DisabledAction, #True)
    DisableGadget(Button_SaveAction, #True)
    DisableGadget(Button_RemoveAction, #True)
    DisableGadget(ButtonImage_MoveUp, #True)
    DisableGadget(ButtonImage_MoveDown, #True)
    DisableGadget(ButtonImage_ShowCommandHelp, #True)
    DisableGadget(Button_DisableAction, #True)
    
    ; Remove panel action
    RemoveGadgetItem(Panel_Options, 1)
    
    ; Return
    ProcedureReturn
  Else
    DisableGadget(String_ActionName, #False)
    DisableGadget(Editor_ActionDescription, #False)
    DisableGadget(Editor_ActionConditions, #False)
    DisableGadget(String_ActionVariable, #False)
    DisableGadget(Checkbox_ContinueOnError, #False)
    DisableGadget(Checkbox_DisabledAction, #False)
    DisableGadget(Button_SaveAction, #False)
    DisableGadget(Button_RemoveAction, #False)
    DisableGadget(ButtonImage_MoveUp, #False)
    DisableGadget(ButtonImage_MoveDown, #False)
    DisableGadget(ButtonImage_ShowCommandHelp, #False)
    DisableGadget(Button_DisableAction, #False)
    
    SetGadgetState(Checkbox_ContinueOnError, #PB_Checkbox_Unchecked)
    SetGadgetState(Checkbox_DisabledAction, #PB_Checkbox_Unchecked)
  EndIf
  
  ; Clear list for parameter gadgets
  ClearList(ActionParameterGadgets())
  
  ; Bind gadget event if input change triggered
  BindGadgetEvent(String_ActionCommand, @EditorInputHandler(), #PB_EventType_Change)
  BindGadgetEvent(String_ActionName, @EditorInputHandler(), #PB_EventType_Change)
  BindGadgetEvent(Editor_ActionDescription, @EditorInputHandler(), #PB_EventType_Change)
  BindGadgetEvent(Editor_ActionConditions, @EditorInputHandler(), #PB_EventType_Change)
  BindGadgetEvent(String_ActionVariable, @EditorInputHandler(), #PB_EventType_Change)
  
  ; Tab - Options
  If DatabaseQuery(1, "SELECT Command, Name, Description, Disabled, ContinueOnError, VariableName, Conditions FROM Actions WHERE ID="+ID+" LIMIT 1")
    While NextDatabaseRow(1)
      Command = GetDatabaseString(1, 0)
      Name = GetDatabaseString(1, 1)
      Description = GetDatabaseString(1, 2)
      Disabled = GetDatabaseLong(1, 3)
      ContinueOnError = GetDatabaseLong(1, 4)
      VariableName = GetDatabaseString(1, 5)
      Conditions = GetDatabaseString(1, 6)
    Wend
    
    SetGadgetText(String_ActionCommand, Command)
    SetGadgetText(String_ActionName, Name)
    SetGadgetText(Editor_ActionDescription, Description)
    SetGadgetText(Editor_ActionConditions, Conditions)
    SetGadgetText(String_ActionVariable, VariableName)
    
    If ContinueOnError = 1
      SetGadgetState(Checkbox_ContinueOnError, #PB_Checkbox_Checked)
    EndIf
    
    If Disabled = 1
      SetGadgetState(Checkbox_DisabledAction, #PB_Checkbox_Checked)
      SetGadgetText(Button_DisableAction, "Enable")
    Else
      SetGadgetText(Button_DisableAction, "Disable")
    EndIf
    
    FinishDatabaseQuery(1) 
  Else
    MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
  EndIf
  
  ; Tab - Action
  RemoveGadgetItem(Panel_Options, 1)
  
  If DatabaseQuery(0, "Select C.Name As Command, P.Parameter, P.Description, P.Control, P.Required, P.Type FROM Parameters As P LEFT JOIN Commands As C ON P.Command = C.ID WHERE C.Command LIKE '"+Command+"' ORDER BY SortIndex ASC")
    OpenGadgetList(Panel_Options)
    AddGadgetItem(Panel_Options, -1, "Action")
    ActionScrollAreaGadget = ScrollAreaGadget(#PB_Any, 0, 0, 325, 705, 300, 705)
    
    Protected Count.i = 1
    Protected GadgetPosY.i = 18
    Protected GadgetDefaultHeight.i = 25
    
    While NextDatabaseRow(0)
      ; Result
      Protected ParameterName.s = GetDatabaseString(0, 1)
      Protected ParameterValue.s = GetActionParameterValueByID(ID, ParameterName)
      Protected ParameterDescription.s = GetDatabaseString(0, 2)
      Protected ParameterRequired.i = GetDatabaseLong(0, 4)
      Protected ParameterType.s = GetDatabaseString(0, 5)
      Protected ControlType.s = GetDatabaseString(0, 3)
      Protected NamePrefix.s = ""
      
      ; Debug
      Debug "[Debug: Action Editor Renderer] Current action gadget count: " + Count
      
      ; Counting
      If Count <> 1
        GadgetPosY = GadgetPosY + 70
      EndIf
      
      ; Prefix
      If ParameterRequired
        NamePrefix = "* "
      EndIf
      
      ; Input
      If Command = "#CustomScript"
        Define TextGadget = TextGadget(#PB_Any, 20, GadgetPosY, 270, GadgetDefaultHeight, NamePrefix+ParameterName+" ("+ParameterType+") (Read-Only)"+":")
        Define InputGadget = EditorGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight), 270, 320, #PB_Editor_ReadOnly)
        Define EditorButtonGadget = ButtonGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight + 330), 270, 25, "Open Editor")
        GadgetToolTip(InputGadget, ParameterDescription)
        SetGadgetText(InputGadget, ParameterValue)
        BindGadgetEvent(EditorButtonGadget, @LoadScriptInEditor())
        ;DisableGadget(InputGadget, #True)
        ScriptEditorGadget = InputGadget
        
      ElseIf ControlType = "Checkbox"
        Debug "[Debug: Action Editor Renderer] Control is checkbox."
        Define TextGadget = TextGadget(#PB_Any, 20, GadgetPosY, 270, GadgetDefaultHeight, NamePrefix+ParameterName+" ("+ParameterType+")"+":")
        Define InputGadget = CheckBoxGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight), 270, GadgetDefaultHeight, "Enabled")
        GadgetToolTip(InputGadget, ParameterDescription)
        
        If ParameterValue = "1" Or ParameterValue = "Enabled"
          SetGadgetState(InputGadget, #PB_Checkbox_Checked)
        Else
          SetGadgetState(InputGadget, #PB_Checkbox_Unchecked)
        EndIf
      Else
        Define TextGadget = TextGadget(#PB_Any, 20, GadgetPosY, 270, GadgetDefaultHeight, NamePrefix+ParameterName+" ("+ParameterType+")"+":")
        Define InputGadget = StringGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight), 270, GadgetDefaultHeight, ParameterValue)
        GadgetToolTip(InputGadget, ParameterDescription)
      EndIf
      
      ; Add to the parameter gadget list
      AddElement(ActionParameterGadgets())
      ActionParameterGadgets()\ActionID = ID
      ActionParameterGadgets()\Parameter = ParameterName
      ActionParameterGadgets()\TitleGadget = TextGadget
      ActionParameterGadgets()\InputGadget = InputGadget
      ActionParameterGadgets()\ControlType = ControlType
      
      ; Bind gadget event if input change triggered
      BindGadgetEvent(InputGadget, @EditorInputHandler(), #PB_EventType_Change)
      
      Count + 1
    Wend
    
    ; Set scrollbar area
    SetGadgetAttribute(ActionScrollAreaGadget, #PB_ScrollArea_InnerHeight, (Count * 70) + (GadgetHeight(Panel_Options) / 2))
    
    CloseGadgetList()
    FinishDatabaseQuery(0)
  Else
    MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure UpdateDatabaseStepCount()
  CheckDatabaseUpdate(1, "UPDATE Actions SET Step = (SELECT new_step FROM (SELECT ID, ROW_NUMBER() OVER (ORDER BY Step) AS new_step FROM Actions WHERE DeploymentType = 'Installation') AS ordered WHERE ordered.ID = Actions.ID) WHERE DeploymentType = 'Installation';")
  CheckDatabaseUpdate(1, "UPDATE Actions SET Step = (SELECT new_step FROM (SELECT ID, ROW_NUMBER() OVER (ORDER BY Step) AS new_step FROM Actions WHERE DeploymentType = 'Uninstall') AS ordered WHERE ordered.ID = Actions.ID) WHERE DeploymentType = 'Uninstall';")  
  CheckDatabaseUpdate(1, "UPDATE Actions SET Step = (SELECT new_step FROM (SELECT ID, ROW_NUMBER() OVER (ORDER BY Step) AS new_step FROM Actions WHERE DeploymentType = 'Repair') AS ordered WHERE ordered.ID = Actions.ID) WHERE DeploymentType = 'Repair';")   
EndProcedure

Procedure SaveAction(EventType)
  Protected ActionName.s = GetGadgetText(String_ActionName)
  Protected ActionDescription.s = GetGadgetText(Editor_ActionDescription)
  Protected ActionConditions.s = GetGadgetText(Editor_ActionConditions)
  Protected ActionVariable.s = GetGadgetText(String_ActionVariable)
  Protected ActionContinueOnError.i = GetGadgetState(Checkbox_ContinueOnError)
  Protected ActionDisabled.i = GetGadgetState(Checkbox_DisabledAction)
  
  If SelectedActionID = 0
    Debug "[Debug: Save Action Handler] No action selected for saving!"
    ProcedureReturn
  EndIf
    
  ;Debug "ContinueOnError: "+ActionContinueOnError
  ;Debug "Disabled: "+ActionDisabled
  
  ; Save the parameters
  Debug "Parameter gadget list size is: "+ListSize(ActionParameterGadgets())
  ForEach ActionParameterGadgets()
    Define GadgetValue.s = ""
    
    ; Get value by control type
    Select ActionParameterGadgets()\ControlType
        
      Case "Checkbox"
        If GetGadgetState(ActionParameterGadgets()\InputGadget) = #PB_Checkbox_Checked
          GadgetValue = "1"
        Else
          GadgetValue = "0"
        EndIf
        
      Default
        GadgetValue = GetGadgetText(ActionParameterGadgets()\InputGadget)
        
    EndSelect
    
    ; Debug
    Debug "[Debug: Save Action Handler] Update action with ID "+ActionParameterGadgets()\ActionID+" And parameter "+ActionParameterGadgets()\Parameter+" With the value: "+GadgetValue
    
    ; Delete the old values
    SetDatabaseLong(1, 0, ActionParameterGadgets()\ActionID)
    SetDatabaseString(1, 1, ActionParameterGadgets()\Parameter)
    CheckDatabaseUpdate(1, "DELETE FROM Actions_Values WHERE Action=? AND Parameter=?")
    
    ; Insert new value
    SetDatabaseLong(1, 0, ActionParameterGadgets()\ActionID)
    SetDatabaseString(1, 1, ActionParameterGadgets()\Parameter)
    SetDatabaseString(1, 2, GadgetValue)
    CheckDatabaseUpdate(1, "INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (?, ?, ?)")
  Next
  
  ; Save the options
  If IsDatabase(1) And SelectedActionID <> 0
    CheckDatabaseUpdate(1, "UPDATE Actions SET Name='"+ActionName+"',Description='"+ActionDescription+"',Conditions='"+ActionConditions+"',VariableName='"+ActionVariable+"',Disabled="+ActionDisabled+",ContinueOnError="+ActionContinueOnError+" WHERE ID="+SelectedActionID)
  EndIf
  
  ; Update step counter
  Debug "[Debug: Save Action Handler] Update step counter"
  UpdateDatabaseStepCount()
  
  ; Finish
  FinishDatabaseQuery(1) 
  RefreshProject(0)
  TreeSequence_ExpandedByID(SelectedActionID)
  UnsavedChange = #False
EndProcedure

Procedure AddAction(EventType)
  Protected LastStep.i = 0, NextStep.i = 0
  Protected NewActionName.s = "Your new action"
  Protected Command.s = GetGadgetText(ListView_Commands)
  Debug "[Debug: Add Action Handler] Adding new action for: "+Command
  
  ; Retrieve last step
  SetDatabaseString(1, 0, CurrentDeploymentType)
  If DatabaseQuery(1, "SELECT MAX(Step) FROM Actions WHERE DeploymentType=? LIMIT 1")
    If FirstDatabaseRow(1)
      LastStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Add Action Handler] Last step count: "+LastStep
    EndIf
  EndIf
  
  ; Calc next step
  NextStep = LastStep + 1
  
  ; Create action name/title based on command
  NewActionName = ReplaceString(Command, "-", " ")
  NewActionName = LCase(NewActionName)
  NewActionName = "New " + NewActionName
  
  ; Add new action in the database
  If IsDatabase(1) And Command <> ""
    Protected Query.s = "INSERT INTO Actions (Step, Command, Name, Disabled, ContinueOnError, DeploymentType) VALUES ("+NextStep+", '"+Command+"', '"+NewActionName+"', 0, 0, '"+CurrentDeploymentType+"')"
    Debug "[Debug: Add Action Handler] Update table: "+Query
    CheckDatabaseUpdate(1, Query)
  EndIf
  
  ; Finish query and refresh project view
  FinishDatabaseQuery(1)
  RefreshProject(0)
  
EndProcedure

Procedure AddAction_CustomScript(EventType)
  Protected LastStep.i = 0, NextStep.i = 0
  Protected Command.s = "#CustomScript"
  
  ; Retrieve last step
  SetDatabaseString(1, 0, CurrentDeploymentType)
  If DatabaseQuery(1, "SELECT MAX(Step) FROM Actions WHERE DeploymentType=? LIMIT 1")
    If FirstDatabaseRow(1)
      LastStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Add Action Custom Script Handler] Last step count: "+LastStep
    EndIf
  EndIf
  
  ; Calc next step
  NextStep = LastStep + 1
  
  ; Add new action
  If IsDatabase(1) And Command <> ""
    Protected Query.s = "INSERT INTO Actions (Step, Command, Name, Disabled, ContinueOnError, DeploymentType) VALUES ("+NextStep+", '"+Command+"', 'My custom PowerShell script', 0, 0, '"+CurrentDeploymentType+"')"
    Debug "[Debug: Add Action Custom Script Handler] Update table: "+Query
    CheckDatabaseUpdate(1, Query)
  EndIf
  
  FinishDatabaseQuery(1)
  RefreshProject(0)
EndProcedure

Procedure RemoveAction(EventType)
  ; Ask first
  Define Result = MessageRequester("Remove", "Are you sure you want to delete this action?", #PB_MessageRequester_YesNo | #PB_MessageRequester_Warning)
  If Result = #PB_MessageRequester_No
    ProcedureReturn
  EndIf
  
  ; Remove action
  If IsDatabase(1)
    Protected RemoveActionQuery.s = "DELETE FROM Actions WHERE ID="+SelectedActionID
    Debug "[Debug: Remove Action Handler] Remove entry in table: "+RemoveActionQuery
    CheckDatabaseUpdate(1, RemoveActionQuery)
    
    Protected RemoveValuesQuery.s = "DELETE FROM Actions_Values WHERE Action="+SelectedActionID
    Debug "[Debug: Remove Action Handler] Remove entry in table: "+RemoveValuesQuery
    CheckDatabaseUpdate(1, RemoveValuesQuery)
    
    Debug "[Debug: Remove Action Handler] Update step counter"
    UpdateDatabaseStepCount()
  EndIf
  
  FinishDatabaseQuery(1)
  RefreshProject(0)
EndProcedure

Procedure DisableAction(EventType)
  ; Toggle Checkbox
  If GetGadgetState(Checkbox_DisabledAction) = #PB_Checkbox_Unchecked
    SetGadgetState(Checkbox_DisabledAction, #PB_Checkbox_Checked)
    SetGadgetText(Button_DisableAction, "Enable")
  Else
    SetGadgetState(Checkbox_DisabledAction, #PB_Checkbox_Unchecked)
    SetGadgetText(Button_DisableAction, "Disable")
  EndIf
  
  ; Save and refresh
  SaveAction(0)
  RefreshProject(0)
  
  ; Expand
  TreeSequence_ExpandedByID(SelectedActionID)
EndProcedure

Procedure MoveActionUp(EventType)
  Protected CurrentStep.i = -1
  
  ; Get current step
  If DatabaseQuery(1, "SELECT Step FROM Actions WHERE ID="+SelectedActionID)
    If FirstDatabaseRow(1)
      CurrentStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Move Action Handler] Current step is: "+CurrentStep
    EndIf
  EndIf
  
  ; Return if 
  If CurrentStep <= 1
    Debug "[Debug: Move Action Handler] The action is already the first one."
    ProcedureReturn
  EndIf
  
  ; Fix overlapping step
  Protected OverlappingStep = CurrentStep-1
  ;Debug "Overlapping step would be: "+OverlappingStep
  ;Debug "Fix query for overlapping: UPDATE Actions SET Step="+CurrentStep+" WHERE Step="+OverlappingStep
  
  SetDatabaseLong(1, 0, CurrentStep)
  SetDatabaseLong(1, 1, OverlappingStep)
  SetDatabaseString(1, 2, CurrentDeploymentType)
  If DatabaseQuery(1, "UPDATE Actions SET Step=? WHERE Step=? AND DeploymentType=?")
    If FirstDatabaseRow(1)
      Debug "[Debug: Move Action Handler] Fixed overlapping step."
    EndIf
  EndIf
  
  ; Update the selected one
  SetDatabaseLong(1, 0, CurrentStep-1)
  SetDatabaseLong(1, 1, SelectedActionID)
  CheckDatabaseUpdate(1, "UPDATE Actions SET Step=? WHERE ID=?")
  FinishDatabaseQuery(1)
  
  RefreshProject(0)
  TreeSequence_ExpandedByID(SelectedActionID)
EndProcedure

Procedure MoveActionDown(EventType)
  Protected CurrentStep.i = -1
  Protected LastStep.i = -1
  
  ; Get last step
  If DatabaseQuery(1, "SELECT MAX(Step) FROM Actions WHERE DeploymentType='"+CurrentDeploymentType+"'")
    If FirstDatabaseRow(1)
      LastStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Move Action Handler] Last step is: "+LastStep
    EndIf
  EndIf
  
  ; Get current step
  If DatabaseQuery(1, "SELECT Step FROM Actions WHERE ID="+SelectedActionID)
    If FirstDatabaseRow(1)
      CurrentStep = GetDatabaseLong(1, 0)
      Debug "[Debug: Move Action Handler] Current step is: "+CurrentStep
    EndIf
  EndIf
  
  ; Return if 
  If CurrentStep >= LastStep
    Debug "[Debug: Move Action Handler] The action is already the last one."
    ProcedureReturn
  EndIf

  ; Fix overlapping step
  Protected OverlappingStep = CurrentStep+1
  ;Debug "Overlapping step would be: "+OverlappingStep
  ;Debug "Fix query for overlapping: UPDATE Actions SET Step="+CurrentStep+" WHERE Step="+OverlappingStep
  
  SetDatabaseLong(1, 0, CurrentStep)
  SetDatabaseLong(1, 1, OverlappingStep)
  SetDatabaseString(1, 2, CurrentDeploymentType)
  If DatabaseQuery(1, "UPDATE Actions SET Step=? WHERE Step=? AND DeploymentType=?")
    If FirstDatabaseRow(1)
      Debug "[Debug: Move Action Handler] Fixed overlapping step."
    EndIf
  EndIf
  
  ; Update the selected one
  SetDatabaseLong(1, 0, CurrentStep+1)
  SetDatabaseLong(1, 1, SelectedActionID)
  CheckDatabaseUpdate(1, "UPDATE Actions SET Step=? WHERE ID=?")
  FinishDatabaseQuery(1)
  
  RefreshProject(0)
  TreeSequence_ExpandedByID(SelectedActionID)
EndProcedure

Procedure StartDeploymentWithPSADT(DeploymentType.s = "Install")
  Protected SilentSwitch.s = ""
  
  If GetGadgetState(Checkbox_SilentMode) = #PB_Checkbox_Checked
    Debug "[Debug: PSADT Deployment Starter] Running PSADT deployment in silent mode: "+DeploymentType
    SilentSwitch = " -DeployMode Silent"
  EndIf

  ShellExecute_(0, "RunAS", Chr(34)+Project_FolderPath+"\Invoke-AppDeployToolkit.exe"+Chr(34), "-DeploymentType '"+DeploymentType+"'"+SilentSwitch, Project_FolderPath, #SW_SHOWNORMAL)
  ;RunProgram(Chr(34)+Project_FolderPath+"\Invoke-AppDeployToolkit.exe"+Chr(34), "-DeploymentType "+DeploymentType+SilentSwitch, Project_FolderPath)
EndProcedure

Procedure StartInstallation(EventType)
  StartDeploymentWithPSADT("Install")
EndProcedure

Procedure StartUninstallation(EventType)
  StartDeploymentWithPSADT("Uninstall")
EndProcedure

Procedure StartRepair(EventType)
  StartDeploymentWithPSADT("Repair")
EndProcedure

Procedure StartPowerShell(EventType)
  ShellExecute_(0, "RunAS", "powershell.exe", "-NoExit -Command " + Chr(34) + "Set-Location '"+Project_FolderPath+"'", Project_FolderPath, #SW_SHOWNORMAL)
EndProcedure

Procedure StartHelp(EventType)
  RunProgram("powershell.exe", "-ExecutionPolicy ByPass -File " + Chr(34) + GetCurrentDirectory() + "Scripts\PSADT4-HelpConsole.ps1", GetCurrentDirectory())
EndProcedure

Procedure StartPowerShellEditor(EventType)
  ShellExecute_(0, "RunAS", "powershell_ise.exe", Chr(34) + Project_DeploymentFile + Chr(34), "", #SW_SHOWNORMAL)
EndProcedure

Procedure StartNotepadPlusPlus(EventType)
  RunProgram("notepad++.exe", Chr(34) + Project_DeploymentFile + Chr(34), "")
EndProcedure

Procedure StartVSCode(EventType)
  RunProgram("code", Chr(34) + Project_DeploymentFile + Chr(34), "")
EndProcedure

Procedure.s BuildScript(DeploymentType.s = "Installation")
  
  Protected ScriptBuilder.s = ""
  Protected PSADT_Spacing.s = Space(4)
  Protected BlockOpened.i = 0 ; <- NEU
  
  ; Read database
  If OpenDatabase(1, Project_Database, "", "")
    If DatabaseQuery(1, "SELECT Action,Command,Parameter,Value,Description,ContinueOnError,VariableName,Conditions FROM View_Sequence WHERE Disabled=0 AND DeploymentType='"+DeploymentType+"' ORDER BY Step ASC")
      Protected LastAction.i = 0
      
      While NextDatabaseRow(1)
        Protected CurrentAction.i = GetDatabaseLong(1, 0)
        Protected Command.s = GetDatabaseString(1, 1)
        Protected Parameter.s = GetDatabaseString(1, 2)
        Protected Value.s = GetDatabaseString(1, 3)
        Protected Description.s = GetDatabaseString(1, 4)
        Protected ContinueOnError.i = GetDatabaseLong(1, 5)
        Protected VariableName.s = GetDatabaseString(1, 6)
        Protected Conditions.s = GetDatabaseString(1, 7)
        Protected EnclosingCharacter.s = Chr(34)
        Protected Prefix_Variable.s = ""
        Protected ParameterType.s = ""
        
        ; Enclosing character by type
        ParameterType = ParameterTypeByCommand(Command, Parameter)
        Debug "[Debug: PSADT Script Builder] "+Command+" > "+Parameter+" = "+ParameterType
        
        If Not FindString(ParameterType, "String")
          EnclosingCharacter = ""
        EndIf
        
        If FindString(ParameterType, "Guid")
          EnclosingCharacter = Chr(34)
        EndIf
          
        ; Description
        Description = RemoveString(Description, #CRLF$)
        
        Select Command
          Case "#CustomScript"
            Debug "[Debug: PSADT Script Builder] Found custom script in the action list!"
            Value = ReplaceString(Value, #CRLF$, #CRLF$ + PSADT_Spacing)
            ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "# " + Description + #CRLF$ + PSADT_Spacing + Value
            LastAction = GetDatabaseLong(1, 0)
            Continue
        EndSelect
        
        If CurrentAction <> LastAction Or (CurrentAction = 0 And Parameter = "")
          If Description <> ""
            ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "# " + Description
          EndIf
          
          If Conditions <> ""
            If BlockOpened
              ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "}"
            EndIf
            ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "if ("+Conditions+") { "
            PSADT_Spacing = Space(6)
            BlockOpened = 1
          Else
            If BlockOpened
              ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "}"
              BlockOpened = 0
            EndIf
            PSADT_Spacing = Space(4)
          EndIf
          
          If Trim(VariableName) <> ""
            Prefix_Variable = VariableName + " = "
          EndIf
        
          ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + Prefix_Variable + Command
          
          If ContinueOnError
            ScriptBuilder = ScriptBuilder+" -ErrorAction SilentlyContinue"
          EndIf
        EndIf
        
        ; Parameters
        If Value <> ""
          Select ParameterType
            Case "SwitchParameter"
              If Value = "1"
                ScriptBuilder = ScriptBuilder+" -"+Parameter
              EndIf
            Default
              ScriptBuilder = ScriptBuilder+" -"+Parameter+" "+EnclosingCharacter+Value+EnclosingCharacter
          EndSelect
        EndIf
        
        LastAction = GetDatabaseLong(1, 0)
      Wend
      
      ; Schliessenden Block am Ende ergänzen
      If BlockOpened
        ScriptBuilder = ScriptBuilder + #CRLF$ + Space(4) + "}"
        BlockOpened = 0
      EndIf
      
      FinishDatabaseQuery(1) 
    Else
      MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
    EndIf
  Else
    MessageRequester("Database Error", "Can't open the database: " + Project_Database, #PB_MessageRequester_Error)
  EndIf
  
  ProcedureReturn ScriptBuilder
EndProcedure

Procedure GenerateDeploymentTemplateFile()
  If FileSize(Project_DeploymentFile + ".template") = -1
    Debug "[Debug: Deployment Template File Copier] Template file is missing!"
    CopyFile(PSADT_TemplateFile, Project_DeploymentFile + ".template")
  EndIf
EndProcedure

Procedure GenerateDeploymentFile()

  ; Create parts
  StatusBarText(0, 0, "Building scripts...")
  Protected InstallationPart.s = BuildScript("Installation")
  Protected UninstallPart.s = BuildScript("Uninstall")
  Protected RepairPart.s = BuildScript("Repair")
  
  ; Check if project deployment file template exists and create if needed
  GenerateDeploymentTemplateFile()
  
  ; Set the project template file as base
  PSADT_TemplateFile = Project_DeploymentFile + ".template"
  
  ; Create empty deployment file
  StatusBarText(0, 0, "Reset deployment file...")
  Protected NewDeploymentFile = CreateFile(#PB_Any, Project_DeploymentFile) 
  CloseFile(NewDeploymentFile)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generating deployment file...")
  Protected FileIn, FileOut, sLine.s
  FileIn = ReadFile(#PB_Any, PSADT_TemplateFile, #PB_File_SharedRead)
  If FileIn
    Debug "[Debug: Deployment File Generator] Generating new deployment file: "+Project_DeploymentFile
    FileOut = OpenFile(#PB_Any, Project_DeploymentFile)
    
    If FileOut
      ; Read each line from the input file, replace text, and write to the output file
      While Not Eof(FileIn)
        sLine = ReadString(FileIn)
        sLine = ReplaceString(sLine, "<ProductPublisher>", GetProjectSetting("App_Vendor"))
        sLine = ReplaceString(sLine, "<ProductName>", GetProjectSetting("App_Name"))
        sLine = ReplaceString(sLine, "<ProductVersion>", GetProjectSetting("App_Version"))
        sLine = ReplaceString(sLine, "<ProductLanguage>", GetProjectSetting("App_Language"))
        sLine = ReplaceString(sLine, "<ProductArchitecture>", GetProjectSetting("App_Architecture"))
        sLine = ReplaceString(sLine, "<ScriptDate>", FormatDate("%yyyy-%mm-%dd", Date()))
        sLine = ReplaceString(sLine, "<ScriptAuthor>", GetProjectSetting("App_Author"))
        sLine = ReplaceString(sLine, "<InstallationPart>", InstallationPart)
        sLine = ReplaceString(sLine, "<UninstallPart>", UninstallPart)
        sLine = ReplaceString(sLine, "<RepairPart>", RepairPart)
        
        WriteString(FileOut, sLine + #CRLF$)
      Wend
      CloseFile(FileOut)
    Else
      MessageRequester("Error", "Can't write new deployment file.", #PB_MessageRequester_Error)
    EndIf
    CloseFile(FileIn)
  EndIf
EndProcedure

Procedure CreateWinGetProject(EventType)
  Protected SelectedWinGetYaml.s, WinGetManifest_FilePath.s
  Protected TargetArchitecture.s = GetGadgetText(Combo_TargetArchitecture)
  Protected Format, Line.s
  Protected FoundArchitecture = #False, FoundInstallerUrl = #False, FoundSilentSwitch = #False
  Protected Identifier.s = "", Version.s = "", InstallerUrl.s = "", Silent.s = ""
  
  ; Find installer url and silent switch in manifest 
  SelectedWinGetYaml = GetGadgetItemText(WinGetImport_ListIcon, GetGadgetState(WinGetImport_ListIcon), 4)
  WinGetManifest_FilePath = WinGet_ManifestTempFolder + "\" + SelectedWinGetYaml
  
  If ReadFile(0, WinGetManifest_FilePath)
    Format = ReadStringFormat(0)
    While Eof(0) = 0
      ; Read current line
      Line = ReadString(0, Format)
      
      ; Search and extract informations
      If FindString(Line, "PackageIdentifier:")
        Identifier = ReplaceString(Line, "PackageIdentifier:", "")
        Identifier = Trim(Identifier)
        Identifier = RemoveString(RemoveString(Identifier, Chr(34)), Chr(39))
        
        Debug "[Debug: YAML] Package Identifier: " + Identifier
      EndIf
   
      If FindString(Line, "PackageVersion:")
        Version = ReplaceString(Line, "PackageVersion:", "")
        Version = Trim(Version)
        Version = RemoveString(RemoveString(Version, Chr(34)), Chr(39))
        
        Debug "[Debug: YAML] Version: " + Version
      EndIf
      
      If FindString(Line, "InstallerType:") And FindString(Line, "nullsoft")
        Silent = "/S"
        FoundSilentSwitch = #True
        
        Debug "[Debug: YAML] Installer is Nullsoft"
      EndIf
      
      If FindString(Line, "Silent:") And FoundArchitecture = #False And FoundInstallerUrl = #False
        Silent = ReplaceString(Line, "Silent:", "")
        Silent = LTrim(Silent)
        
        Debug "[Debug: YAML] Found main silent switch for all architectures: " + Silent
        FoundSilentSwitch = #True
      EndIf
      
      If FindString(Line, "- Architecture:") And FindString(Line, TargetArchitecture) And FoundInstallerUrl = #False
        Debug "[Debug: YAML] Found target architecture: " + TargetArchitecture
        FoundArchitecture = #True
      EndIf
      
      If FoundArchitecture And FoundInstallerUrl = #False And FindString(Line, "InstallerUrl:")
        InstallerUrl = ReplaceString(Line, "InstallerUrl:", "")
        InstallerUrl = Trim(InstallerUrl)
        
        Debug "[Debug: YAML] Installer url: " + InstallerUrl
        FoundInstallerUrl = #True
      EndIf
      
      If FoundArchitecture And FoundInstallerUrl And FoundSilentSwitch = #False And FindString(Line, "Silent:")
        Silent = ReplaceString(Line, "Silent:", "")
        Silent = LTrim(Silent)
        
        Debug "[Debug: YAML] Found silent switch: " + Silent
        FoundSilentSwitch = #True
      EndIf
      
      If FoundArchitecture And FoundInstallerUrl And FoundSilentSwitch
        Debug "[Debug: YAML] Found all details to create the project."
        Break
      EndIf
    Wend
    
    ; Check for prereqs
    If FoundArchitecture = #False Or FoundInstallerUrl = #False
      ProcedureReturn MessageRequester("WinGet Import", "The WinGet package cannot be imported. The installer with "+TargetArchitecture+" support is missing.", #PB_MessageRequester_Error | #PB_MessageRequester_Ok)
    EndIf
    
    CloseFile(0)
  Else
    MessageRequester("Information", "Couldn't open the manifest file: " + WinGetManifest_FilePath)
  EndIf
  
  ; Open progress window
  CloseWinGetImportWindow(0)

  ; Create the new project by the user
  If Not CreateNewProject(0)
    ProcedureReturn
  EndIf
  
  ; Download installer file
  OpenProgressWindow()
  
  ; Start the asynchronous download to a file
  Download_Url = InstallerUrl
  Download_OutputFile = Project_FolderPath + "Files\Installer." + GetExtensionPart(Download_Url)
  WinGet_Identifier = Identifier
  WinGet_Version = Version
  WinGet_SilentSwitch = Silent
  CreateThread(@DownloadInstallerFile(), 0)
  
  ;RedrawWindow_(MainWindow, #Null, #Null, #RDW_INVALIDATE | #RDW_UPDATENOW)
  UpdateWindow_(MainWindow)
  
  ; App name and vendor
  Protected CompanyCharactersLength.i = FindString(WinGet_Identifier, ".", 0)

  ; Update project settings
  If IsDatabase(1)
    UpdateProjectSettings("Database_Version", "1.0.4")
    UpdateProjectSettings("Project_Name", "WinGet Project")
    UpdateProjectSettings("App_Version", WinGet_Version)
    UpdateProjectSettings("App_Vendor", StringField(WinGet_Identifier, 1, "."))
    UpdateProjectSettings("App_Architecture", TargetArchitecture)
    UpdateProjectSettings("App_Language", "EN")
    UpdateProjectSettings("App_Author", "WinGet Import by Deployment Editor")
    UpdateProjectSettings("App_Name", ReplaceString(Right(WinGet_Identifier, Len(WinGet_Identifier) - CompanyCharactersLength), ".", " "))
    FinishDatabaseQuery(1)
  EndIf
  
  ; Load new values
  LoadProjectSettings()
  
  ; Generate new deployment file
  GenerateDeploymentFile()
  
  ; Update Main Window
  SetGadgetText(Text_ProjectName, GetProjectSetting("Project_Name"))
  
  ; Message to the user
  MessageRequester("WinGet Import", "Once the installer has been successfully downloaded in the background, the action/parameter for silent installation will be added automatically.", #PB_MessageRequester_Info | #PB_MessageRequester_Ok)
  
EndProcedure

Procedure GenerateAndStartInstallation(EventType)
  ; Save first the current action options
  SaveAction(0)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generate deployment file...")
  GenerateDeploymentFile()
  Delay(1000)
  
  ; Starting deployment installation
  StatusBarText(0, 0, "Starting Deployment > Installation")
  StartInstallation(0)
EndProcedure

Procedure GenerateAndStartInstallationSandbox(EventType)
  ; Save first the current action options
  SaveAction(0)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generate deployment file...")
  GenerateDeploymentFile()
  Delay(1000)
  
  ; Starting deployment installation
  StatusBarText(0, 0, "Starting Deployment > Installation (in Windows Sandbox)")
  
  ; Copy template file to project folder
  CopyFile(PSADT_SandboxTemplate, Project_FolderPath + "Windows Sandbox.wsb")
  
  ; Update Windows Sandbox file
  Protected FileIn, FileOut, sLine.s
  FileIn = ReadFile(#PB_Any, PSADT_SandboxTemplate, #PB_File_SharedRead)
  If FileIn
    FileOut = CreateFile(#PB_Any, Project_FolderPath + "Windows Sandbox.wsb")
    
    If FileOut
      ; Read each line from the input file, replace text, and write to the output file
      While Not Eof(FileIn)
        sLine = ReadString(FileIn)
        sLine = ReplaceString(sLine, "$(ProjectPath)", Project_FolderPath)
        
        WriteString(FileOut, sLine + #CRLF$)
      Wend
      CloseFile(FileOut)
    Else
      MessageRequester("Error", "Can't write new Windows Sandbox configuration file.", #PB_MessageRequester_Error)
    EndIf
    CloseFile(FileIn)
  EndIf
  
  ; Run Windows Sandbox
  RunProgram(Project_FolderPath + "Windows Sandbox.wsb")
EndProcedure

Procedure GenerateAndStartUninstall(EventType)
  ; Save first the current action options
  SaveAction(0)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generate deployment file...")
  GenerateDeploymentFile()
  Delay(1000)
  
  ; Starting deployment installation
  StatusBarText(0, 0, "Starting Deployment > Uninstall")
  StartUninstallation(0)
EndProcedure

Procedure GenerateAndStartRepair(EventType)
  ; Save first the current action options
  SaveAction(0)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generate deployment file...")
  GenerateDeploymentFile()
  Delay(1000)
  
  ; Starting deployment installation
  StatusBarText(0, 0, "Starting Deployment > Repair")
  StartRepair(0)
EndProcedure

Procedure OnlyGenerateDeploymentFile(EventType)
  ; Save first the current action options
  SaveAction(0)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generate deployment file...")
  GenerateDeploymentFile()
  
  ; Final
  StatusBarText(0, 0, "Done.")
  ;TreeSequence_AllExpanded()
  
  ; Message to the user
  MessageRequester("Information", "The new deployment file has been successfully generated: " + Project_DeploymentFile, #PB_MessageRequester_Info)

EndProcedure

Procedure SwitchDeploymentType()
  Protected DeploymentType.s = GetGadgetText(Combo_DeploymentType)
  CurrentDeploymentType = DeploymentType
  RefreshProject(0)
EndProcedure

Procedure DownloadIntuneWinAppUtil(Event)
  
  Protected DestinationPath.s = GetTemporaryDirectory() + "IntuneWinAppUtil.exe"
  Protected DownloadFileSize = 0

  If ReceiveHTTPFile("https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe", DestinationPath)
    Debug "[Debug: IntuneWinAppUtil] IntuneWinAppUtil file size is: " + FileSize(DestinationPath)
    DownloadFileSize = FileSize(DestinationPath)
    
    If DownloadFileSize >= 50000 And DownloadFileSize <= 100000
      Debug "[Debug: IntuneWinAppUtil] Downloaded IntuneWinAppUtil successfully: " + DestinationPath
      IntuneWinAppUtil = DestinationPath
    Else
      Debug "[Debug: IntuneWinAppUtil] Failed to download IntuneWinAppUtil - Wrong file size."
      MessageRequester("Error", "Unfortunately, something went wrong during the download of IntuneWinAppUtil. Please check your internet connection or place the file here: " + DestinationPath, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
    EndIf
  Else
    Debug "[Debug: IntuneWinAppUtil] Failed to download IntuneWinAppUtil - Network issue"
    MessageRequester("Download Error", "Unfortunately, something went wrong during the download of IntuneWinAppUtil. Please check your internet connection or place the file here: " + DestinationPath, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
  
EndProcedure

Procedure LoadPlugins(Event)
  Delay(200)
  Debug "Examine plugin folder..."
  
  ; Reset
  ClearGadgetItems(ListView_Plugins)
  ResetList(EditorPlugins()) 
  
  If ExamineDirectory(0, PluginDirectory, "*.*")
    While NextDirectoryEntry(0)
      Protected FileName$ = DirectoryEntryName(0)
      
      If DirectoryEntryType(0) = #PB_DirectoryEntry_Directory
        If FindString(FileName$, "Enabled-", 0)
          Protected PluginFolder.s = PluginDirectory + "\" + FileName$
          Protected PluginName.s = ReplaceString(FileName$, "Enabled-", "")
          
          Debug "[Debug: Plugin Loader] Found plugin which is enabled by folder name: "+PluginName+" ("+PluginFolder+")"
          AddGadgetItem(ListView_Plugins, -1, PluginName)
          
          ; Read preference file
          Protected PluginPreference.s = PluginFolder+"\Plugin.ini"
          Debug PluginPreference
          OpenPreferences(PluginPreference)
          PreferenceGroup("Details")
          
          ; Add plugin to the list
          AddElement(EditorPlugins())
          EditorPlugins()\ID = PluginName
          EditorPlugins()\Name = ReadPreferenceString("Name", "")
          EditorPlugins()\Version = ReadPreferenceString("Version", "")
          EditorPlugins()\Date = ReadPreferenceString("Date", "")
          EditorPlugins()\Description = ReadPreferenceString("Description", "")
          EditorPlugins()\Author = ReadPreferenceString("Author", "")
          EditorPlugins()\Website = ReadPreferenceString("Website", "")
          PreferenceGroup("Script")
          EditorPlugins()\Path = PluginFolder
          EditorPlugins()\File = ReadPreferenceString("File", "")
          EditorPlugins()\Parameter = ReadPreferenceString("Parameter", "")
          EditorPlugins()\UnloadProject = ReadPreferenceString("UnloadProject", "")
        Else
          Continue
        EndIf
      EndIf
    Wend
  Else
    MessageRequester("Error","Can't examine this directory: "+GetGadgetText(0),0)
  EndIf
  
  DisableGadget(Button_RunPlugin, #False)
EndProcedure

Procedure RenderPluginDetails(Event)
  Protected SelectedPluginID.s = GetGadgetText(ListView_Plugins)
  Debug SelectedPluginID
  
  ForEach EditorPlugins()
    If EditorPlugins()\ID = SelectedPluginID
      Debug "[Debug: Plugin Loader] Found details about the plugin!"
      SetGadgetText(Text_PluginDetails, EditorPlugins()\Name+#CRLF$+EditorPlugins()\Description+#CRLF$+#CRLF$+"Developed by "+EditorPlugins()\Author+#CRLF$+"Version: "+EditorPlugins()\Version)
      
      ; Stop loop
      Break
    EndIf
  Next
EndProcedure

Procedure CloseProject(Event)
  CloseDatabase(1)
  DisableMainWindowGadgets(#True)
  ClearGadgetItems(Tree_Sequence)
EndProcedure

Procedure RunPlugin(EventType)
  Protected SelectedPluginID.s = GetGadgetText(ListView_Plugins)
  
  If Trim(SelectedPluginID) = ""
    ProcedureReturn MessageRequester("Error", "First select the plugin you wish to run from the list view.", #PB_MessageRequester_Warning | #PB_MessageRequester_Ok)  
  EndIf
  
  ForEach EditorPlugins()
    If EditorPlugins()\ID = SelectedPluginID
      Debug "[Debug: Plugin Loader] Found the plugin! Lets run it."
      Protected FilePath.s = EditorPlugins()\Path + "\" + EditorPlugins()\File
      
      ; Check if the project need to be closed
      If EditorPlugins()\UnloadProject = "Yes"
        Protected UnloadProjectRequester = MessageRequester("Unload Project to Continue", "To continue and execute the plugin script, the project must be unloaded from the Deployment Editor. Do you want to save all actions and continue?", #PB_MessageRequester_YesNo | #PB_MessageRequester_Info)
        
        If UnloadProjectRequester = #PB_MessageRequester_Yes
          CloseProject(0)
          ClosePluginWindow(0)
          HideWindow(MainWindow, #True)
          ShowNewProjectWindow(0)
        Else
          ProcedureReturn 0
        EndIf
      EndIf

      ; Run process
      Protected PowerShell_Parameter.s = "-ExecutionPolicy ByPass -File "+Chr(34)+FilePath+Chr(34)+" -ProjectPath "+Chr(34)+Project_FolderPath
      Debug "[Debug: Plugin Loader] PowerShell parameter: " + PowerShell_Parameter
      RunProgram("powershell.exe", PowerShell_Parameter, GetPathPart(FilePath))
      
      ; Stop loop
      Break
    EndIf
  Next
EndProcedure

Procedure LoadPreviewInEditor(EventType)
  Debug "[Debug: Deployment File Preview] Loading Monaco Editor..."
  
  Protected Script$ = BuildScript(CurrentDeploymentType)
  HideWindow(ScriptEditorWindow, #False)
  
  WebViewExecuteScript(WebView_ScriptEditor, "myMonacoEditor.setValue('"+EscapeString(Script$)+"')")
  WebViewExecuteScript(WebView_ScriptEditor, "myMonacoEditor.updateOptions({ readOnly: true });")
  BindWebViewCallback(WebView_ScriptEditor, "updateScript", @EmptyCallback())
EndProcedure

Procedure ShowPluginWindow(EventType)
  If IsWindow(PluginWindow)
    HideWindow(PluginWindow, #False)
    SetActiveWindow(PluginWindow)
  Else
    OpenPluginWindow()
  EndIf
  
  ; Load Plugins
  Debug "[Debug: Plugin Loader] Loading plugins..."
  CreateThread(@LoadPlugins(), 0)
EndProcedure

Procedure ClosePluginWindow(EventType)
  If IsWindow(PluginWindow)
    HideWindow(PluginWindow, #True)
    HideWindow(MainWindow, #False)
    SetActiveWindow(MainWindow)
  EndIf
EndProcedure

Procedure EndApplication(EventType)
  
  ; Save Project if Database is open
  If IsDatabase(1)
    Define SaveDialog = MessageRequester("Deployment Editor", "Would you like to save the current project before closing?", #PB_MessageRequester_YesNoCancel | #PB_MessageRequester_Info)
    If SaveDialog = #PB_MessageRequester_Yes
      SaveAction(0)
      Delay(500)
    ElseIf SaveDialog = #PB_MessageRequester_Cancel
      ProcedureReturn 0
    EndIf
  EndIf
  
  ; Close PSADT database
  If IsDatabase(0)
    Debug "[Debug: Application Ending Process] Closing Commands Database: " + PSADT_Database
    CloseDatabase(0)
  EndIf
  
  ; Close Project database
  If IsDatabase(1)
    Debug "[Debug: Application Ending Process] Closing Project Database: " + Project_Database
    CloseDatabase(1)
  EndIf
  
  ; Shutdown application
  End
  
EndProcedure

;------------------------------------------------------------------------------------
;- Main Loop
;------------------------------------------------------------------------------------
ShowMainWindow()
LoadUI()
ShowNewProjectWindow(0)
ShowSoftwareReadMe()

; Script Editor (preload)
OpenScriptEditorWindow()
Delay(400)
SetGadgetText(WebView_ScriptEditor, "file://" + GetCurrentDirectory() + "Web/MonacoEditor/editor.html")

; New Thread: Download IntuneWinAppUtil
Debug "[Debug: IntuneWinAppUtil] Installing IntuneWinAppUtil from GitHub."
CreateThread(@DownloadIntuneWinAppUtil(), 0)

;------------------------------------------------------------------------------------
;- Event Loop
;------------------------------------------------------------------------------------
Repeat
  Event = WaitWindowEvent()

  Select EventWindow()
      
    ;- [Main Window]
    Case MainWindow
      
      ; Event = Close Window
      If Event = #PB_Event_CloseWindow
        EndApplication(0)
        
      ; Event = Menu
      ElseIf Event = #PB_Event_Menu
        Select EventMenu()
          
          ; Keyboard
          Case #Keyboard_Shortcut_Run : GenerateAndStartInstallation(0)
          Case #Keyboard_Shortcut_Save : SaveAction(0)
            
          ; Menu
          Case #MenuItem_Open : OpenOtherProject(0)
          Case #MenuItem_New : ShowNewProjectWindow(0)
          Case #MenuItem_Save : SaveAction(0)
          Case #MenuItem_Reload : RefreshProject(0)
          Case #MenuItem_Quit : EndApplication(0)
          Case #MenuItem_ShowLogs : ShowSoftwareLogFolder(EventMenu())
          Case #MenuItem_AboutApp : ShowAboutWindow(0)
          Case #MenuItem_PSADT_OnlineDocumentation : ShowOnlineDocumentation(0)
          Case #MenuItem_OpenWithISE : StartPowerShellEditor(0)
          Case #MenuItem_OpenWithNotepadPlusPlus : StartNotepadPlusPlus(0)
          Case #MenuItem_OpenWithVSCode : StartVSCode(0)
          Case #MenuItem_ShowProjectFolder : ShowProjectFolder(0)
          Case #MenuItem_ShowFilesFolder : ShowFilesFolder(0)
          Case #MenuItem_ShowSupportFilesFolder : ShowSupportFilesFolder(0)
          Case #MenuItem_PSADT_OnlineVariablesOverview : ShowOnlineVariablesOverview(0)
          Case #MenuItem_RunInstallation : GenerateAndStartInstallation(0)
          Case #MenuItem_RunInstallationSandbox : GenerateAndStartInstallationSandbox(0)
          Case #MenuItem_RunRemoteMachine : NotAvailableFeatureMessage(0)
          Case #MenuItem_RunUninstall : GenerateAndStartUninstall(0)
          Case #MenuItem_RunRepair : GenerateAndStartRepair(0)
          Case #MenuItem_GenerateExecutablesList : GenerateExecutablesList(0)
          Case #MenuItem_CreateIntunePackage : CreateIntunePackage(0)
          Case #MenuItem_ShowPlugins : ShowPluginWindow(0)
        EndSelect
        
      ; Gadget = Combo for Deployment Type
      ElseIf Event = #PB_Event_Gadget And EventGadget() = Combo_DeploymentType
        Select EventType()
          Case #PB_EventType_Change
            SwitchDeploymentType()
        EndSelect
        
      ; Gadget = Tree for Sequence View
      ElseIf Event = #PB_Event_Gadget And EventGadget() = Tree_Sequence
        Select EventType()
          Case #PB_EventType_LeftClick
            RenderActionEditor()
          Case #PB_EventType_RightClick       : Debug "Click with right mouse button"
          Case #PB_EventType_LeftDoubleClick  : Debug "Double-click with left mouse button"
          Case #PB_EventType_RightDoubleClick : Debug "Double-click with right mouse button"
        EndSelect
        
        
      ; GadgetDrop = Tree for Sequence View 
      ElseIf Event = #PB_Event_GadgetDrop And EventGadget() = Tree_Sequence
        TreeSequence_DropHandler(EventDropFiles())
        
      ; Gadget = Command Search
      ElseIf Event = #PB_Event_Gadget And EventGadget() = String_CommandSearch
        Select EventType()
          Case #PB_EventType_Change
            FilterCommands(GetGadgetText(String_CommandSearch))
        EndSelect
        
      Else
        ; Default events
        MainWindow_Events(Event)
      EndIf
      
    ;- [New Project Window]
    Case NewProjectWindow
      If Event = #PB_Event_CloseWindow
        CloseNewProjectWindow(0)
      Else
        NewProjectWindow_Events(Event)
      EndIf
      
    ;- [Project Settings Window]
    Case ProjectSettingsWindow
      If Event = #PB_Event_CloseWindow
        CloseProjectSettingsWindow(0)
      Else
        ProjectSettingsWindow_Events(Event)
      EndIf
      
    ;- [Plugin Window]
    Case PluginWindow
      If Event = #PB_Event_CloseWindow
        ClosePluginWindow(0)
      ElseIf Event = #PB_Event_Gadget And EventGadget() = ListView_Plugins
        Select EventType()
          Case #PB_EventType_LeftClick
            RenderPluginDetails(Event)
        EndSelect
      Else
        PluginWindow_Events(Event)
      EndIf
      
    ;- [About Window]
    Case AboutWindow
      If Event = #PB_Event_CloseWindow
        CloseAboutWindow(0)
      Else
        AboutWindow_Events(Event)
      EndIf
      
    ;- [WinGet Import Window]
    Case WinGetImportWindow
      If Event = #PB_Event_CloseWindow
        CloseWinGetImportWindow(0)
      ElseIf Event = #PB_Event_Menu
          Select EventMenu()
            Case #WinGetImport_Enter : SearchWinGetPackage(Event)
          EndSelect
      Else
        WinGetImportWindow_Events(Event)
      EndIf
      
    ;- [Progress Window]
    Case ProgressWindow
      If Event = #PB_Event_CloseWindow
        CloseProgressWindow(0)
      Else
        ProgressWindow_Events(Event)
      EndIf
      
     ;- [Import Exe Window]
    Case ImportExeWindow
      If Event = #PB_Event_CloseWindow
        CloseImportExeWindow(0)
      Else
        ImportExeWindow_Events(Event)
      EndIf
      
     ;- [Import Msi Window]
    Case ImportMsiWindow
      If Event = #PB_Event_CloseWindow
        CloseImportMsiWindow(0)
      Else
        ImportMsiWindow_Events(Event)
      EndIf
      
     ;- [Script Editor Window]
    Case ScriptEditorWindow
      If Event = #PB_Event_CloseWindow
        CloseScriptEditorWindow(0)
      Else
        ScriptEditorWindow_Events(Event)
      EndIf
      
  EndSelect
  
Until Quit = #True
; IDE Options = PureBasic 6.30 beta 6 (Windows - x64)
; CursorPosition = 2935
; FirstLine = 2902
; Folding = --------------------
; EnableXP
; DPIAware