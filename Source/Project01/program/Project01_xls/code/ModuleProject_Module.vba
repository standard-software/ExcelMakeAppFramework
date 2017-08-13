'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Project Module
'ObjectName:    ModuleProject
'--------------------------------------------------
'Version:       2015/07/29
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'■プロジェクト設定
'--------------------------------------------------
Public Const Project_Name As String = _
    "Project01"

Public Const Project_AppID As String = _
    "StandardSoftware.ExcelMakeAppFramework." + Project_Name
    
Public Const Project_ScriptFileName As String = _
    Project_Name + ".vbs"

Public Const Project_ProgramFolderName As String = _
    "program"

Public Const Project_StartMenuFolderName As String = _
    "Excel MakeApp"
Public Const Project_ShortcutFileName As String = _
    Project_Name
    
Public Const Project_FormMainTitle As String = _
    Project_Name
    
Public Const Project_FormCreateAppShortcut_Title As String = _
    Project_Name

Public Const Project_MainIconFileName As String = _
    "FormMainIcon.ico"
Public Const Project_MainIconIndex As Long = _
    0

'--------------------------------------------------
'■バージョン情報
'--------------------------------------------------
Public Const Project_VersionNumberText As String = _
    "1.0.0"
Public Const Project_VersionDialogWindowTitle As String = _
    Project_Name + " のバージョン情報"
    
Public Const Project_VersionDialogInstruction As String = _
    "バージョン情報"
    
Public Const Project_VersionDialogContent As String = _
    Project_Name + vbCrLf + _
    "Version " + Project_VersionNumberText

