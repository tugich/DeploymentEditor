;-- Start
UseSQLiteDatabase()
#Separator = ","
DatabaseFile$ = GetCurrentDirectory()+"PSADT.sqlite"

;-- Procedures
Procedure CheckDatabaseUpdate(Database, Query$)
  Debug Query$
  
  Result = DatabaseUpdate(Database, Query$)
  If Result = 0
    Debug DatabaseError()
  EndIf
  
  ProcedureReturn Result
EndProcedure
 
Procedure.s RemoveSpecialCharacters(String$)
  String$ = RemoveString(String$, Chr(34))
  String$ = RemoveString(String$, Chr(39))
  
  ProcedureReturn String$
EndProcedure

;-- Commands Table
File = ReadFile(#PB_Any, "PSADT4_Commands.csv")
If File
  
  If OpenDatabase(0, DatabaseFile$, "", "")
    Count = 0
    
    ; Truncate table and set autoincrement
    CheckDatabaseUpdate(0, "DELETE FROM Commands;")
    CheckDatabaseUpdate(0, "UPDATE sqlite_sequence SET seq = 0 WHERE name= 'Commands'")
    
    While Not Eof(File)
      Count = Count + 1
      Line$ = ReadString(File)
      
      ; Ignore first line
      If Count = 1
        Debug "Ignore the first line."
        Continue
      EndIf
      
      ; Define fields
      ID$ = StringField(Line$, 1, #Separator)
      Name$ = ReplaceString(StringField(Line$, 2, #Separator), "-", " ")
      Command$ = StringField(Line$, 2, #Separator)
      Description$ = RemoveSpecialCharacters(StringField(Line$, 3, #Separator))
      
      ; Run SQL query
      SQLQuery$ = "INSERT INTO Commands (ID, Name, Command, Category, Description) VALUES ("+ID$+", "+Name$+", "+Command$+", 'PSADT', '"+Description$+"')"
      CheckDatabaseUpdate(0, SQLQuery$)
    Wend
  Else
    Debug "Can't open database !"
  EndIf
  
  CloseFile(File)
  CloseDatabase(0)
EndIf

;-- Parameters Table
File = ReadFile(#PB_Any, "PSADT4_Parameters.csv")
If File
  
  If OpenDatabase(0, DatabaseFile$, "", "")
    Count = 0
    SortIndex = 1
    LastCommandID$ = ""
    
    ; Truncate table and set autoincrement
    CheckDatabaseUpdate(0, "DELETE FROM Parameters;")
    CheckDatabaseUpdate(0, "UPDATE sqlite_sequence SET seq = 0 WHERE name= 'Parameters'")
    
    While Not Eof(File)
      Count = Count + 1
      Line$ = ReadString(File)
      
      ; Ignore first line
      If Count = 1
        Debug "Ignore the first line."
        Continue
      EndIf
      
      ; Define fields
      CommandID$ = StringField(Line$, 1, #Separator)
      ParameterName$ = StringField(Line$, 2, #Separator)
      ParameterType$ = StringField(Line$, 3, #Separator)
      ParameterRequired$ = RemoveSpecialCharacters(StringField(Line$, 4, #Separator)) 
      ParameterDescription$ = RemoveSpecialCharacters(StringField(Line$, 5, #Separator)) 
      
      ; Ignore default PowerShell parameters:
      If FindString("Verbose Debug ErrorAction WarningAction InformationAction ErrorVariable WarningVariable InformationVariable OutVariable OutBuffer PipelineVariable PassThru Callback", RemoveSpecialCharacters(ParameterName$))
        Continue
      Else
        ;Debug RemoveSpecialCharacters(ParameterName$)
      EndIf
      
      ; Define sort index
      If CommandID$ <> LastCommandID$
        SortIndex = 1
      Else
        SortIndex = SortIndex + 1
      EndIf
      
      ; Define control type
      If FindString(ParameterType$, "SwitchParameter")
        ControlType$ = "Checkbox"
      Else
        ControlType$ = "String"
      EndIf
      
      ; Run SQL query
      SQLQuery$ = "INSERT INTO Parameters (Command, Parameter, Type, Description, Required, Control, SortIndex) VALUES ("+CommandID$+", "+ParameterName$+", "+ParameterType$+", '"+ParameterDescription$+"', "+ParameterRequired$+", '"+ControlType$+"', "+SortIndex+")"
      CheckDatabaseUpdate(0, SQLQuery$)
      
      LastCommandID$ = CommandID$
    Wend
  Else
    Debug "Can't open database !"
  EndIf
  
  ; Add custom entries
  CheckDatabaseUpdate(0, "INSERT INTO Commands (Name, Command, Category, Description) VALUES ('Custom Script', '#CustomScript', 'General', 'Custom scripts in PowerShell')")
  CheckDatabaseUpdate(0, "INSERT INTO Parameters (Command, Parameter, Type, Description, Required, Control, SortIndex) VALUES ((SELECT last_insert_rowid()), 'Script', 'Inline', 'Your custom PowerShell script.', 1, 'Script', 1)")
  CheckDatabaseUpdate(0, "INSERT INTO Parameters (Command, Parameter, Type, Description, Required, Control, SortIndex) VALUES (101, 'Action', 'String', 'Specifies the action to be performed. Available options: Install, Uninstall, Patch, Repair, ActiveSetup.', 1, 'String', 1)")
  
  CloseFile(File)
  CloseDatabase(0)
EndIf

; IDE Options = PureBasic 6.20 (Windows - x64)
; CursorPosition = 127
; FirstLine = 76
; Folding = -
; EnableXP
; DPIAware
; Executable = ..\DeploymentEditor.exe