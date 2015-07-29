'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Project03 Sheet1
'ObjectName:    Sheet1
'--------------------------------------------------
'Version:       2015/07/29
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

    Dim XOffset As Long: XOffset = 14
    Dim XOffsetResizeOn As Long: XOffsetResizeOn = 8
    Dim XOffsetResizeOff As Long: XOffsetResizeOff = 4
    Dim YOffsetTitleBarOn As Long: YOffsetTitleBarOn = 20
    Dim YOffsetTitleBarOff As Long: YOffsetTitleBarOff = 0
    Dim YOffsetResizeOn As Long: YOffsetResizeOn = 8
    Dim YOffsetResizeOff As Long: YOffsetResizeOff = 4

    XOffset = XOffset * (GetDPI / 96)
    XOffsetResizeOn = XOffsetResizeOn * (GetDPI / 96)
    XOffsetResizeOff = XOffsetResizeOff * (GetDPI / 96)
    YOffsetTitleBarOn = YOffsetTitleBarOn * (GetDPI / 96)
    YOffsetTitleBarOff = YOffsetTitleBarOff * (GetDPI / 96)
    YOffsetResizeOn = YOffsetResizeOn * (GetDPI / 96)
    YOffsetResizeOff = YOffsetResizeOff * (GetDPI / 96)

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
        Call MsgBox( _
            Project_VersionDialogInstruction + vbNewLine + _
            Project_VersionDialogContent, _
            vbOKOnly, _
            Project_VersionDialogWindowTitle)
    Case "AppClose"
        ThisWorkbook.Saved = True
        ThisWorkbook.Close
    End Select
End Sub

'--------------------------------------------------
'■プログラム本体
'--------------------------------------------------
'以下にプログラム本体の処理を記述してください
'--------------------------------------------------

