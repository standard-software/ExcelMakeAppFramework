'--------------------------------------------------
'Excel Make App Framework
'Main Module
'
'ObjectName:    ModuleMain
'--------------------------------------------------
'バージョン     2014/12/06
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'■Main
'--------------------------------------------------
Sub testMain()
    Select Case 1
    Case 1
        Call Main("")
    Case 2
        Call Main("Hello" + vbTab + "World")
    End Select
End Sub

Sub Main(ByVal ArgsText As String)
On Error GoTo Err:

    Call Load(FormMain)
    FormMain.Args = ExcludeLastStr(ArgsText, vbTab)
    
    Call FormMain.Initialize( _
        TaskBarButton:=True, _
        TitleBar:=True, _
        SystemMenu:=True, _
        FormIcon:=True, _
        MinimizeButton:=True, _
        MaximizeButton:=True, _
        CloseButton:=True, _
        ResizeFrame:=True, _
        TopMost:=False)
        
    Call FormMain.Show(vbModal)
    Call Unload(FormMain)

    Exit Sub
Err:
    Call MsgBox( _
        CStr(Err.Number) + vbCrLf + _
        Err.Source + vbCrLf + _
        Err.Description)
End Sub

'--------------------------------------------------
'■プロジェクト共通関数
'--------------------------------------------------

'----------------------------------------
'◆プロジェクトファイルパス
'----------------------------------------
Function Project_MainFolderPath() As String
    Project_MainFolderPath = _
        fso.GetParentFolderName(ThisWorkbook.Path)
End Function

Function ProjectScriptFilePath() As String
    ProjectScriptFilePath = PathCombine( _
        Project_MainFolderPath, _
        Project_ScriptFileName)
End Function

Function Project_IniFilePath() As String
    Project_IniFilePath = PathCombine( _
        Project_MainFolderPath, _
        Project_Name + ".ini")
End Function

Function Project_MainIconFilePath() As String
    Project_MainIconFilePath = PathCombine( _
        Project_MainFolderPath, _
        Project_ProgramFolderName, _
        Project_MainIconFileName)
End Function

'----------------------------------------
'◆コマンドファイルパス
'----------------------------------------
Function Project_TaskbarPinCommandExeFilePath() As String
    Project_TaskbarPinCommandExeFilePath = PathCombine( _
        Project_MainFolderPath, _
        Project_ProgramFolderName, _
        "tool", _
        "Win7AppId.exe")
End Function


'----------------------------------------
'◆ショートカットファイルパス
'----------------------------------------
Function Project_ShortcutFilePath_Desktop() As String
    Project_ShortcutFilePath_Desktop = PathCombine( _
        GetSpecialFolderPath(Desktop), _
        Project_ShortcutFileName + ".lnk")
End Function

Function Project_ShortcutFilePath_StartMenu() As String
    Project_ShortcutFilePath_StartMenu = PathCombine( _
        GetSpecialFolderPath(StartMenuProgram), _
        Project_StartMenuFolderName, _
        Project_ShortcutFileName + ".lnk")
End Function

Function ProjectShortcutFileFolderPath_StartMenuGroup() As String
    ProjectShortcutFileFolderPath_StartMenuGroup = PathCombine( _
        GetSpecialFolderPath(StartMenuProgram), _
        Project_StartMenuFolderName)
End Function

Function Project_ShortcutFilePath_SendTo() As String
    Project_ShortcutFilePath_SendTo = PathCombine( _
        GetSpecialFolderPath(SendTo), _
        Project_ShortcutFileName + ".lnk")
End Function

Function Project_ShortcutFilePath_TaskbarPin() As String
    Project_ShortcutFilePath_TaskbarPin = PathCombine( _
        GetSpecialFolderPath(TaskbarPin), _
        Project_ShortcutFileName + ".lnk")
End Function

Function Project_ShortcutFilePath_TaskbarPinCscript() As String
    Project_ShortcutFilePath_TaskbarPinCscript = _
        PathCombine(GetSpecialFolderPath(TaskbarPin), _
            "Microsoft " + ChrW(&HAE) + " Console Based Script Host.lnk")
End Function





