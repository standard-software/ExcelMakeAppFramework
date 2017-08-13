'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Main Module
'ObjectName:    ModuleMain
'--------------------------------------------------
'Version:       2015/07/29
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

    Call SetCurrentProcessExplicitAppUserModelID( _
        StrPtr(Project_AppID))

    ThisWorkbook.OriginalWindowRectBuffer = _
         RectToStr(Form_GetRectPixel(Application))

    Dim WorkAreaRect As Rect
    WorkAreaRect = GetRectWorkArea
    Application.Width = GetRectWidth(WorkAreaRect) \ 3
    Application.Height = GetRectHeight(WorkAreaRect) \ 3
    Call Form_IniReadPosition(Application, _
        Project_IniFilePath, _
        "Form", "Rect", False)

    Application.Visible = True
    Call ApplicationModeOn
    
    Call SetExcelWindowTitle(Project_FormMainTitle)
    
    Call SetWindowStyle(Application.hWnd, _
        TitleBar:=True, _
        SystemMenu:=True, _
        ResizeFrame:=True, _
        MinimizeButton:=True, _
        MaximizeButton:=True)
    
    Call SetWindowIcon(Application.hWnd, _
        Project_MainIconFilePath, Project_MainIconIndex)


    Dim ArgsArray() As String
    ArgsArray = Split(ArgsText, vbTab)
    Dim I As Long
    For I = 0 To ArrayCount(ArgsArray) - 1
        Call MsgBox(ArgsArray(I))
    Next
    
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





