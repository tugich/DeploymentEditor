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
;UseGIFImageDecoder()
UseSQLiteDatabase()

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

;------------------------------------------------------------------------------------
;- Variables, Enumerations and Maps
;------------------------------------------------------------------------------------
Global Event = #Null, Quit = #False
Global MainWindowTitle.s = "Deployment Editor ("+#DE_Version+") - TUGI.CH"
Global DonationUrl.s = "https://www.paypal.com/donate/?hosted_button_id=PXABL8ESQQ4F8"

; Templates
Global Template_PSADT.s = GetCurrentDirectory() + "ThirdParty\PSAppDeployToolkit\"
Global Template_EmptyDatabase.s = GetCurrentDirectory() + "Templates\Deploy-Application.db"

; PSADT
Global PSADT_Database.s = GetCurrentDirectory() + "Databases\PSADT.sqlite"
Global PSADT_TemplateFile.s = GetCurrentDirectory() + "Templates\Deploy-Application.ps1"
Global PSADT_OnlineDocumentation.s = "https://psappdeploytoolkit.com/docs"

; Intune
Global IntuneWinAppUtil.s = ""

; Project
Global Project_FolderPath.s = GetCurrentDirectory() + "Test\"
Global Project_DeploymentFile.s = Project_FolderPath + "Deploy-Application.ps1"
Global Project_Database.s = Project_FolderPath + "Deploy-Application.db"

; Deployment
Global CurrentDeploymentType.s = "Installation"

; Maps
Global NewMap PSADT_Parameters.PSADT_Parameter()
Global NewMap ProjectSettings.ProjectSetting()

; Lists
Global NewList RecentProjects.RecentProject()
; > Add Demo Project
AddElement(RecentProjects())
RecentProjects()\FileName = "Deploy-Application.db"
RecentProjects()\FolderPath = Project_FolderPath

; Action
Global SelectedActionID.i
Global UnsavedChange = #False
Global NewList ActionParameterGadgets.PSADT_Parameter_GadgetGroup()

; WinAPI
Global tvi.TV_ITEM

; Shortcuts
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
  #MenuItem_RunInstallation
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
  Debug "Update query: " + Query$
  
  Result = DatabaseUpdate(Database, Query$)
  If Result = 0
    Debug DatabaseError()
  Else
    Debug "Database update was successfully."
    Debug "Affected rows: " + AffectedDatabaseRows(Database)
  EndIf
  
  ProcedureReturn Result
EndProcedure

;------------------------------------------------------------------------------------
;- Functions
;------------------------------------------------------------------------------------
Procedure ShowSoftwareReadMe()
  CompilerIf #PB_Compiler_Debugger = 0
    MessageRequester("Readme", Readme, #PB_MessageRequester_Info)
  CompilerEndIf
EndProcedure

Procedure ShowMainWindow()
  OpenMainWindow()
  WindowBounds(MainWindow, WindowWidth(MainWindow)-100, WindowHeight(MainWindow), #PB_Ignore, #PB_Ignore)
  BindEvent(#PB_Event_SizeWindow, @ResizeGadgetsMainWindow(), MainWindow)
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

Procedure LoadProjectSettings()
  
  ; Clear first the map with the old values
  ClearMap(ProjectSettings())
  
  ; Read all new settings from the database
  If OpenDatabase(1, Project_Database, "", "")
    Debug "Loaded Project Database successfully: " + PSADT_Database
    
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
    HideWindow(ProjectSettingsWindow, #True)
    SetActiveWindow(MainWindow)
  EndIf
EndProcedure

Procedure UpdateProjectSettingByGadget(SettingName.s, Gadget)
  SetDatabaseString(1, 0, GetGadgetText(Gadget))
  SetDatabaseString(1, 1, SettingName)
  CheckDatabaseUpdate(1, "UPDATE Settings SET Value = ? WHERE Name = ?")
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
    
    Debug "Active window ID is: "+GetActiveWindow()
    Debug "New project window ID is: "+NewProjectWindow
    
    If GetActiveWindow() = NewProjectWindow
      SetActiveWindow(NewProjectWindow)
    Else
      SetActiveWindow(MainWindow)
    EndIf
  EndIf
EndProcedure

Procedure ShowLicensing(EventType)
  RunProgram(GetCurrentDirectory() + "LICENSE.txt", "", "")
EndProcedure

Procedure OpenDonationUrl(EventType)
  RunProgram(DonationUrl, "", "")
EndProcedure

Procedure NotAvailableFeatureMessage(EventType)
  MessageRequester("Feature not available", "This feature is not yet available for this version. Sign up for the newsletter on the website to be kept up to date.", #PB_MessageRequester_Info | #PB_MessageRequester_Ok) 
EndProcedure

Procedure ShowSoftwareLogFolder(EventType)
  Debug "Start Windows Explorer for Software Logs."
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
  RunProgram(PSADT_OnlineDocumentation, "", "")
EndProcedure

Procedure ShowOnlineVariablesOverview(EventType)
  RunProgram(PSADT_OnlineDocumentation + "/variables", "", "")
EndProcedure

Procedure ShowCommandHelp(EventType)
  Protected CurrentCommand.s = GetGadgetText(String_ActionCommand)
  RunProgram(PSADT_OnlineDocumentation + "/reference/functions/"+CurrentCommand, "", "")
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
    Debug "Loaded Project Database successfully: " + Project_Database
    
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
          AddGadgetItem(Tree_Sequence, -1, "No values for parameters found.", 0, 1)
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

Procedure RefreshProject(EventType)
  ClearGadgetItems(Tree_Sequence)
  LoadProjectFile()
  TreeSequence_SetFirstLevelBold()
  ;TreeSequence_AllExpanded()
EndProcedure

Procedure LoadCommandsAndParameters()
  Protected PSADT_Command.s, PSADT_Category.s
  
  If OpenDatabase(0, PSADT_Database, "", "")
    Debug "Loaded PSADT Database successfully: " + PSADT_Database
    
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
  Debug "Triggered LoadUI()"
  
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

  InstallerFile = "Deploy-Application.exe"
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
  
  Debug "Parameters for the wrapper: " + WrapperParameters$
  
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
    Debug "Successfully created the Intune package."
    MessageRequester("Intune Package Creation", "The Intune package has been successfully created: "+PackagePath, #PB_MessageRequester_Info | #PB_MessageRequester_Ok)
  Else
    Debug "Error Intune package creation!"
    MessageRequester("Intune Package Creation", "The Intune package compiler failed.", #PB_MessageRequester_Error | #PB_MessageRequester_Ok)
  EndIf
  
  ; Load project again
  RefreshProject(0)
  
EndProcedure

Procedure OpenFirstRecentProject(EventType)
  
  ; Set Project Folder and File
  SelectElement(RecentProjects(), 0)
  Project_FolderPath.s = RecentProjects()\FolderPath
  Project_DeploymentFile.s = Project_FolderPath + "Deploy-Application.ps1"
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
  DisableGadget(ListView_Commands, #False)
  DisableGadget(Combo_DeploymentType, #False)
  DisableGadget(Tree_Sequence, #False)
  DisableGadget(Button_SaveAction, #False)
  DisableGadget(Button_AddCommand, #False)
  DisableGadget(Button_AddCustomScript, #False)
  
EndProcedure

Procedure OpenOtherProject(EventType)
  
  ; Request file location
  Protected StandardFile$, Pattern$, File$, Pattern
  StandardFile$ = "C:\"
  Pattern$ = "Project Database (*.db)|*.db|All files (*.*)|*.*"
  Pattern = 0
  File$ = OpenFileRequester("Please choose project database file to load", StandardFile$, Pattern$, Pattern)
  
  If File$ = ""
    Debug "Canceled project database selection."
    ProcedureReturn 0
  EndIf
  
  ; Set Project Folder and File
  Project_FolderPath.s = GetPathPart(File$)
  Project_DeploymentFile.s = Project_FolderPath + "Deploy-Application.ps1"
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
  DisableGadget(ListView_Commands, #False)
  DisableGadget(Combo_DeploymentType, #False)
  DisableGadget(Tree_Sequence, #False)
  DisableGadget(Button_SaveAction, #False)
  DisableGadget(Button_AddCommand, #False)
  DisableGadget(Button_AddCustomScript, #False)
  
EndProcedure

Procedure CreateNewProject(EventType)
  
  ; Set destination path for the new project
  Protected InitialPath$, Path$, ExamineFolder
  InitialPath$ = "C:\"
  Path$ = PathRequester("Please choose your path for the new project", InitialPath$)
  
  If Path$
    Debug "Choosen path is: " + Path$
  Else
    Debug "Abort project creation - No folder selected."
    ProcedureReturn 0
  EndIf
  
  ; Ask user for confirmation
  Define Confirmation = MessageRequester("Confirmation", "Please confirm the destination folder first - all files will be overwritten with the default template files: " + Path$, #PB_MessageRequester_YesNoCancel | #PB_MessageRequester_Warning)
  If Confirmation = #PB_MessageRequester_No Or Confirmation = #PB_MessageRequester_Cancel
    ProcedureReturn MessageRequester("Cancelled", "You have canceled the creation of a new project.", #PB_MessageRequester_Ok | #PB_MessageRequester_Info)
  EndIf
  
  ; Copy template folder from PSADT source
  Debug "Copy PSADT framework..."
  Debug "Source Template folder: " + Template_PSADT
  Debug "Destination folder: " + Path$
  CopyDirectory(Template_PSADT, Path$, "", #PB_FileSystem_Recursive | #PB_FileSystem_Force)
  
  ; Copy empty database template to destination folder
  Debug "Copy empty database file..."
  CopyFile(Template_EmptyDatabase, Path$ + "Deploy-Application.db")
  
  ; Set Project Folder and File
  Project_FolderPath.s = Path$
  Project_DeploymentFile.s = Project_FolderPath + "Deploy-Application.ps1"
  Project_Database.s = Project_FolderPath + "Deploy-Application.db"
  
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
  DisableGadget(ListView_Commands, #False)
  DisableGadget(Combo_DeploymentType, #False)
  DisableGadget(Tree_Sequence, #False)
  DisableGadget(Button_SaveAction, #False)
  DisableGadget(Button_AddCommand, #False)
  DisableGadget(Button_AddCustomScript, #False)
  
EndProcedure

Procedure.s GetActionParameterValueByID(ActionID.i, Parameter.s)
  Protected Value.s = ""
  Protected RequestQuery.s = "SELECT Value FROM Actions_Values WHERE Parameter='"+Parameter+"' AND Action="+ActionID
  
  Debug "Query is: " + RequestQuery
  
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
  Debug "Input in Editor > Action triggered."
  UnsavedChange = #True
EndProcedure

Procedure OptionsInputHandler()
  Debug "Input in Editor > Options triggered."
  SaveAction(0)
EndProcedure

Procedure RenderActionEditor(ID.i = -1)
  
  ; Default values
  Protected SelectedItem.i = GetGadgetState(Tree_Sequence)
  Protected Command.s = "", Name.s = "", Description.s = "", VariableName.s = ""
  Protected Disabled.i = 0, ContinueOnError.i = 0
  
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

  Debug "You clicked on item with data ID: " + ID
  SelectedActionID = ID
  
  If ID = 0
    Debug "No action rendering by ID=0."
    SetGadgetText(String_ActionCommand, "")
    SetGadgetText(String_ActionName, "")
    SetGadgetText(String_ActionVariable, "")
    SetGadgetText(Editor_ActionDescription, "")
    SetGadgetState(Checkbox_ContinueOnError, #PB_Checkbox_Unchecked)
    SetGadgetState(Checkbox_DisabledAction, #PB_Checkbox_Unchecked)
    
    ; Disable gadgets
    DisableGadget(String_ActionName, #True)
    DisableGadget(Editor_ActionDescription, #True)
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
  BindGadgetEvent(String_ActionVariable, @EditorInputHandler(), #PB_EventType_Change)
  
  ; Tab - Options
  If DatabaseQuery(1, "SELECT Command, Name, Description, Disabled, ContinueOnError, VariableName FROM Actions WHERE ID="+ID+" LIMIT 1")
    While NextDatabaseRow(1)
      Command = GetDatabaseString(1, 0)
      Name = GetDatabaseString(1, 1)
      Description = GetDatabaseString(1, 2)
      Disabled = GetDatabaseLong(1, 3)
      ContinueOnError = GetDatabaseLong(1, 4)
      VariableName = GetDatabaseString(1, 5)
    Wend
    
    SetGadgetText(String_ActionCommand, Command)
    SetGadgetText(String_ActionName, Name)
    SetGadgetText(Editor_ActionDescription, Description)
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
      Debug "Current action gadget count: " + Count
      
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
        Define TextGadget = TextGadget(#PB_Any, 20, GadgetPosY, 290, GadgetDefaultHeight, NamePrefix+ParameterName+" ("+ParameterType+")"+":")
        Define InputGadget = ScintillaGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight), 290, 640, 0)
        GadgetToolTip(InputGadget, ParameterDescription)
        
        ; Set value for Scintilla gadget
        Define *Text=UTF8(ParameterValue)
        ScintillaSendMessage(InputGadget, #SCI_SETTEXT, 0, *Text)
        FreeMemory(*Text)
        
      ElseIf ControlType = "Checkbox"
        Debug "Control is checkbox."
        Define TextGadget = TextGadget(#PB_Any, 20, GadgetPosY, 290, GadgetDefaultHeight, NamePrefix+ParameterName+" ("+ParameterType+")"+":")
        Define InputGadget = CheckBoxGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight), 290, GadgetDefaultHeight, "Enabled")
        GadgetToolTip(InputGadget, ParameterDescription)
        
        If ParameterValue = "1" Or ParameterValue = "Enabled"
          SetGadgetState(InputGadget, #PB_Checkbox_Checked)
        Else
          SetGadgetState(InputGadget, #PB_Checkbox_Unchecked)
        EndIf
      Else
        Define TextGadget = TextGadget(#PB_Any, 20, GadgetPosY, 290, GadgetDefaultHeight, NamePrefix+ParameterName+" ("+ParameterType+")"+":")
        Define InputGadget = StringGadget(#PB_Any, 20, (GadgetPosY + GadgetDefaultHeight), 290, GadgetDefaultHeight, ParameterValue)
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
    
    CloseGadgetList()
    FinishDatabaseQuery(0)
  Else
    MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure SaveAction(EventType)
  Protected ActionName.s = GetGadgetText(String_ActionName)
  Protected ActionDescription.s = GetGadgetText(Editor_ActionDescription)
  Protected ActionVariable.s = GetGadgetText(String_ActionVariable)
  Protected ActionContinueOnError.i = GetGadgetState(Checkbox_ContinueOnError)
  Protected ActionDisabled.i = GetGadgetState(Checkbox_DisabledAction)
  
  If SelectedActionID = 0
    Debug "No action selected for saving!"
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
    Debug "Update action with ID "+ActionParameterGadgets()\ActionID+" And parameter "+ActionParameterGadgets()\Parameter+" With the value: "+GadgetValue
    
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
    CheckDatabaseUpdate(1, "UPDATE Actions SET Name='"+ActionName+"',Description='"+ActionDescription+"',VariableName='"+ActionVariable+"',Disabled="+ActionDisabled+",ContinueOnError="+ActionContinueOnError+" WHERE ID="+SelectedActionID)
  EndIf
  
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
  Debug "Adding new action for: "+Command
  
  ; Retrieve last step
  SetDatabaseString(1, 0, CurrentDeploymentType)
  If DatabaseQuery(1, "SELECT MAX(Step) FROM Actions WHERE DeploymentType=? LIMIT 1")
    If FirstDatabaseRow(1)
      LastStep = GetDatabaseLong(1, 0)
      Debug "Last step count: "+LastStep
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
    Debug "Update table: "+Query
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
      Debug "Last step count: "+LastStep
    EndIf
  EndIf
  
  ; Calc next step
  NextStep = LastStep + 1
  
  ; Add new action
  If IsDatabase(1) And Command <> ""
    Protected Query.s = "INSERT INTO Actions (Step, Command, Name, Disabled, ContinueOnError, DeploymentType) VALUES ("+NextStep+", '"+Command+"', 'My custom PowerShell script', 0, 0, '"+CurrentDeploymentType+"')"
    Debug "Update table: "+Query
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
    Debug "Remove entry in table: "+RemoveActionQuery
    CheckDatabaseUpdate(1, RemoveActionQuery)
    
    Protected RemoveValuesQuery.s = "DELETE FROM Actions_Values WHERE Action="+SelectedActionID
    Debug "Remove entry in table: "+RemoveValuesQuery
    CheckDatabaseUpdate(1, RemoveValuesQuery)
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
      Debug "Current step is: "+CurrentStep
    EndIf
  EndIf
  
  ; Return if 
  If CurrentStep <= 1
    Debug "The action is already the first one."
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
      Debug "Fixed overlapping step."
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
      Debug "Last step is: "+LastStep
    EndIf
  EndIf
  
  ; Get current step
  If DatabaseQuery(1, "SELECT Step FROM Actions WHERE ID="+SelectedActionID)
    If FirstDatabaseRow(1)
      CurrentStep = GetDatabaseLong(1, 0)
      Debug "Current step is: "+CurrentStep
    EndIf
  EndIf
  
  ; Return if 
  If CurrentStep >= LastStep
    Debug "The action is already the last one."
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
      Debug "Fixed overlapping step."
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
    Debug "Running PSADT deployment in silent mode: "+DeploymentType
    SilentSwitch = " -DeployMode 'Silent'"
  EndIf

  ShellExecute_(0, "RunAS", Chr(34)+Project_FolderPath+"\Deploy-Application.exe"+Chr(34), "-DeploymentType '"+DeploymentType+"'"+SilentSwitch, "", #SW_SHOWNORMAL)
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
  ShellExecute_(0, "RunAS", "powershell.exe", "", Project_FolderPath, #SW_SHOWNORMAL)
EndProcedure

Procedure StartHelp(EventType)
  ShellExecute_(0, "RunAS", "powershell.exe", "-ExecutionPolicy ByPass -File " + Chr(34) + Project_FolderPath + "\AppDeployToolkit\AppDeployToolkitHelp.ps1" + Chr(34) + "", "", #SW_SHOWNORMAL)
EndProcedure

Procedure StartPowerShellEditor(EventType)
  ShellExecute_(0, "RunAS", "powershell_ise.exe", Chr(34) + Project_DeploymentFile + Chr(34), "", #SW_SHOWNORMAL)
EndProcedure

Procedure StartNotepadPlusPlus(EventType)
  RunProgram("notepad++.exe", Chr(34) + Project_DeploymentFile + Chr(34), "")
EndProcedure

Procedure.s BuildScript(DeploymentType.s = "Installation")
  
  Protected ScriptBuilder.s = ""
  Protected PSADT_Spacing.s = Space(8)
  
  ; Read database
  If OpenDatabase(1, Project_Database, "", "")
    If DatabaseQuery(1, "SELECT Action,Command,Parameter,Value,Description,ContinueOnError,VariableName FROM View_Sequence WHERE Disabled=0 AND DeploymentType='"+DeploymentType+"' ORDER BY Step ASC")
      Protected LastAction.i = 0
      
      While NextDatabaseRow(1)
        Protected CurrentAction.i = GetDatabaseLong(1, 0)
        Protected Command.s = GetDatabaseString(1, 1)
        Protected Parameter.s = GetDatabaseString(1, 2)
        Protected Value.s = GetDatabaseString(1, 3)
        Protected Description.s = GetDatabaseString(1, 4)
        Protected ContinueOnError.i = GetDatabaseLong(1, 5)
        Protected VariableName.s = GetDatabaseString(1, 6)
        Protected EnclosingCharacter.s = Chr(34)
        Protected Prefix_Variable.s = ""
        Protected ParameterType.s = ""
        
        ; Enclosing character by type
        ParameterType = ParameterTypeByCommand(Command, Parameter)
        If ParameterType <> "String"
          EnclosingCharacter = ""
        EndIf
        
        ; Fix Description if any new lines are defined
        Description = RemoveString(Description, #CRLF$)
        
        ; Special commands handling
        Select Command
            
          Case "#CustomScript"
            Debug "Found custom script in the action list!"
            Value = ReplaceString(Value, #CRLF$, #CRLF$ + PSADT_Spacing)
            ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "# " + Description + #CRLF$ + PSADT_Spacing + Value
            LastAction = GetDatabaseLong(1, 0)
            Continue
            
        EndSelect
        
        ; Build PowerShell command with parameters
        If CurrentAction <> LastAction Or (CurrentAction = 0 And Parameter = "")
          If Description <> ""
            ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + "# " + Description
          EndIf
          
          If Trim(VariableName) <> ""
            Prefix_Variable = VariableName + " = "
          EndIf
        
          ScriptBuilder = ScriptBuilder + #CRLF$ + PSADT_Spacing + Prefix_Variable + Command
          
          If ContinueOnError
            ScriptBuilder = ScriptBuilder+" -ContinueOnError 1"
          EndIf
        EndIf

        If Value <> ""
          Select ParameterType
            Case "SwitchParameter"
              If Value = "1"
                ScriptBuilder = ScriptBuilder+" -"+Parameter
              EndIf
              
            Default
              ScriptBuilder = ScriptBuilder+" -"+Parameter+" "+EnclosingCharacter+Value+EnclosingCharacter+""
          EndSelect
        EndIf
        
        LastAction = GetDatabaseLong(1, 0)
      Wend
      
      FinishDatabaseQuery(1) 
    Else
      MessageRequester("Database Error", "Can't execute the query: " + DatabaseError(), #PB_MessageRequester_Error)
    EndIf
  Else
    MessageRequester("Database Error", "Can't open the database: " + Project_Database, #PB_MessageRequester_Error)
  EndIf
  
  ProcedureReturn ScriptBuilder
EndProcedure

Procedure GenerateDeploymentFile()

  ; Create parts
  StatusBarText(0, 0, "Building scripts...")
  Protected InstallationPart.s = BuildScript("Installation")
  Protected UninstallPart.s = BuildScript("Uninstall")
  Protected RepairPart.s = BuildScript("Repair")
  
  ; Create empty deployment file
  StatusBarText(0, 0, "Reset deployment file...")
  Protected NewDeploymentFile = CreateFile(#PB_Any, Project_DeploymentFile) 
  CloseFile(NewDeploymentFile)
  
  ; Generate deployment file
  StatusBarText(0, 0, "Generating deployment file...")
  Protected FileIn, FileOut, sLine.s
  FileIn = ReadFile(#PB_Any, PSADT_TemplateFile, #PB_File_SharedRead)
  If FileIn
    Debug "Generating new deployment file: "+Project_DeploymentFile
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
        sLine = ReplaceString(sLine, "<ScriptDate>", FormatDate("%dd/%mm/%yyyy", Date()))
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
    Debug "IntuneWinAppUtil file size is: " + FileSize(DestinationPath)
    DownloadFileSize = FileSize(DestinationPath)
    
    If DownloadFileSize >= 50000 And DownloadFileSize <= 100000
      Debug "Downloaded IntuneWinAppUtil successfully: " + DestinationPath
      IntuneWinAppUtil = DestinationPath
    Else
      Debug "Failed to download IntuneWinAppUtil - Wrong file size."
      MessageRequester("Error", "Unfortunately, something went wrong during the download of IntuneWinAppUtil. Please check your internet connection.", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
    EndIf
  Else
    Debug "Failed to download IntuneWinAppUtil - Network issue"
    MessageRequester("Download Error", "Unfortunately, something went wrong during the download of IntuneWinAppUtil. Please check your internet connection.", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
  
EndProcedure

Procedure LoadPlugins(Event)
  Debug "Examine plugin folder..."
  
  If ExamineDirectory(0, GetCurrentDirectory() + "\Plugins", "*.*")
    While NextDirectoryEntry(0)
      Protected FileName$ = DirectoryEntryName(0)
      
      If DirectoryEntryType(0) = #PB_DirectoryEntry_Directory
        If FindString(FileName$, "Enabled-", 0)
          Debug "Found plugin which is enabled by folder name: " + ReplaceString(FileName$, "Enabled-", "")          
        EndIf
      EndIf
    Wend
  Else
    MessageRequester("Error","Can't examine this directory: "+GetGadgetText(0),0)
  EndIf
EndProcedure

Procedure ShowPluginWindow(EventType)
  If IsWindow(PluginWindow)
    HideWindow(PluginWindow, #False)
    SetActiveWindow(PluginWindow)
  Else
    OpenPluginWindow()
  EndIf
  
  ; Load Plugins
  Debug "Loading plugins..."
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
    Debug "Closing Commands Database: " + PSADT_Database
    CloseDatabase(0)
  EndIf
  
  ; Close Project database
  If IsDatabase(1)
    Debug "Closing Project Database: " + Project_Database
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

; Download IntuneWinAppUtil
Debug "Downloading IntuneWinAppUtil..."
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
          Case #MenuItem_ShowProjectFolder : ShowProjectFolder(0)
          Case #MenuItem_ShowFilesFolder : ShowFilesFolder(0)
          Case #MenuItem_ShowSupportFilesFolder : ShowSupportFilesFolder(0)
          Case #MenuItem_PSADT_OnlineVariablesOverview : ShowOnlineVariablesOverview(0)
          Case #MenuItem_RunInstallation : GenerateAndStartInstallation(0)
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
      
    ;- [About Window]
    Case PluginWindow
      If Event = #PB_Event_CloseWindow
        ClosePluginWindow(0)
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
      
  EndSelect
  
Until Quit = #True
; IDE Options = PureBasic 6.12 LTS (Windows - x64)
; CursorPosition = 1739
; FirstLine = 293
; Folding = AgAAAAAAAAAAw
; EnableXP
; DPIAware