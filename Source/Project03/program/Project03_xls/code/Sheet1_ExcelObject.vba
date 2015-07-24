'--------------------------------------------------
'Excel Make App Framework
'Project03 Sheet1
'
'ObjectName:    Sheet1
'--------------------------------------------------
Option Explicit

Private Sub ImageMenuButton_Click()
    Dim PopupMenu As CommandBar
    Set PopupMenu = Application.CommandBars.Add(, Position:=msoBarPopup)
    
    Dim MenuItemCreateAppShortcut As CommandBarControl
    Set MenuItemCreateAppShortcut = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemCreateAppShortcut.Caption = "アプリケーションのショートカットを作成..."
    MenuItemCreateAppShortcut.FaceId = 0
    MenuItemCreateAppShortcut.OnAction = PopupMenu_ActionText("CreateAppShortcut")
    
    Dim MenuItemVersionInfo As CommandBarControl
    Set MenuItemVersionInfo = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemVersionInfo.Caption = "バージョン情報"
    MenuItemVersionInfo.FaceId = 0
    MenuItemVersionInfo.OnAction = PopupMenu_ActionText("VersionInfo")
    
    Dim MenuItemAppClose As CommandBarControl
    Set MenuItemAppClose = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemAppClose.BeginGroup = True
    MenuItemAppClose.Caption = "終了"
    MenuItemAppClose.FaceId = 0
    MenuItemAppClose.OnAction = PopupMenu_ActionText("AppClose")
    
    Dim TitleBar As Boolean
    Dim SystemMenu As Boolean
    Dim ResizeFrame As Boolean
    Dim MinimizeButton As Boolean
    Dim MaximizeButton As Boolean
    Call GetWindowStyle(Application.hWnd, _
        TitleBar, SystemMenu, ResizeFrame, _
        MinimizeButton, MaximizeButton)
    Dim TopMost As Boolean
    TopMost = GetWindowTopMost(Application.hWnd)
    
    Select Case PopupMenu_PopupReturn(PopupMenu, _
        ActiveWindow.PointsToScreenPixelsX(0) + _
            PointToPixel(ImageMenuButton.Left) _
            + IIf(ResizeFrame, 8, 4) - 10, _
        ActiveWindow.PointsToScreenPixelsY(0) + _
            PointToPixel(ImageMenuButton.Top + ImageMenuButton.Height) _
            + IIf(ResizeFrame, 8, 4) - 10)
    Case "CreateAppShortcut"
        Call Load(FormCreateAppShortcut)
        Call FormCreateAppShortcut.ShowDialog( _
            Nothing, TopMost)
        Call Unload(FormCreateAppShortcut)
    Case "VersionInfo"
        Call TaskDialogMsgBox( _
            Application.hWnd, _
            ExtractIcon(0, Project_MainIconFilePath, Project_MainIconIndex), _
            Project_VersionDialogWindowTitle, _
            Project_VersionDialogInstruction, _
            Project_VersionDialogContent, _
            TDCBF_OK_BUTTON, _
            True)
    Case "AppClose"
        ThisWorkbook.Saved = True
        ThisWorkbook.Close
    End Select
End Sub
