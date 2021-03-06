'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Main Form
'ObjectName:    FormMain
'--------------------------------------------------
'Version:       2020/04/11
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'■宣言
'--------------------------------------------------
'----------------------------------------
'◆フレームワーク用
'----------------------------------------
Public Args As String

Private FormProperty As New st_vba_FormProperty

Private AnchorMenuButton As New st_vba_ControlAnchor

'----------------------------------------
'◆ユーザー用
'----------------------------------------
'------------------------------
'◇アンカー定義
'------------------------------
'------------------------------
'◇変数定義
'------------------------------

'--------------------------------------------------
'■実装
'--------------------------------------------------

'----------------------------------------
'◆起動・終了
'----------------------------------------

'------------------------------
'◇変数初期化など
'------------------------------
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    Args = ""
    Call IniRead_UserFormInitialize
End Sub

'------------------------------
'◇Mainからの呼び出し
'------------------------------
Public Sub Initialize( _
ByVal TaskBarButton As Boolean, _
ByVal TitleBar As Boolean, _
ByVal SystemMenu As Boolean, _
ByVal FormIcon As Boolean, _
ByVal MinimizeButton As Boolean, _
ByVal MaximizeButton As Boolean, _
ByVal CloseButton As Boolean, _
ByVal ResizeFrame As Boolean, _
ByVal TopMost As Boolean)

    '------------------------------
    '◇フレームワーク初期化処理
    '------------------------------
    With Nothing
        Call FormProperty.InitializeForm(Me)

        Call FormProperty.InitializeProperty( _
            TaskBarButton:=TaskBarButton, _
            TitleBar:=TitleBar, _
            SystemMenu:=SystemMenu, _
            FormIcon:=FormIcon, _
            MinimizeButton:=MinimizeButton, _
            MaximizeButton:=MaximizeButton, _
            CloseButton:=CloseButton, _
            ResizeFrame:=ResizeFrame, _
            TopMost:=TopMost)

        FormProperty.IconPath = Project_MainIconFilePath
        FormProperty.IconIndex = Project_MainIconIndex

        Me.Caption = Project_FormMainTitle
    End With

    '------------------------------
    '◇ユーザー用初期化処理
    '------------------------------
    '以下にユーザー独自の初期化処理を記述してください
    '------------------------------

End Sub

Private Sub UserForm_Activate()
    If FormProperty.Initializing Then
        FormProperty.Initializing = False

        Call SetTaskbarButtonAppID(Project_AppID)

        If FormProperty.Handle = 0 Then
            Call FormProperty.InitializeForm(Me)
            FormProperty.GetWindowsProperty
        Else
            FormProperty.SetWindowsProperty
        End If

        '------------------------------
        '◇メニューボタンを右上端にする
        '------------------------------
        Me.ImageMenuButton.Top = 0
        Me.ImageMenuButton.Left = _
            Me.ImageMenuButton.Parent.InsideWidth - _
            Me.ImageMenuButton.Width + 1

        '------------------------------
        '◇フレームワークアンカー初期化処理
        '------------------------------
        Call AnchorMenuButton.Initialize( _
            Me.ImageMenuButton, _
            HorizonAnchorType.haRight, 2, _
            VerticalAnchorType.vaTop, 0)

        '------------------------------
        '◇ユーザー用アンカー初期化処理
        '------------------------------
        '以下にユーザー独自のアンカー初期化処理を記述してください
        '------------------------------

        Call IniRead_UserFormActivate

        'レイアウトアンカーを動作させる
        Call UserForm_Resize

        Call FormProperty.ForceActiveMouseClick

        '------------------------------
        '◇引数処理
        '------------------------------
        '引数に対する処理を記述してください
        '------------------------------
        Dim ArgsArray() As String
        ArgsArray = Split(Args, vbTab)
        Dim I As Long
        For I = 0 To ArrayCount(ArgsArray) - 1
            Call MsgBox(ArgsArray(I))
        Next
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Select Case CloseMode
    Case 0
        Call IniWrite
    Case 1
    End Select
End Sub

'------------------------------
'◇終了時に呼び出す関数
'------------------------------
'Me.Hide や Call Unload(Me) ではなく
'このFormClose関数を呼び出してください
'Me.Hide や Call Unload(Me) では
'UserForm_QueryCloseイベントが呼び出されず
'Iniファイルへの保存が行われません。
'------------------------------
Private Sub FormClose()
    Dim Cancel As Integer
    Cancel = False
    Call UserForm_QueryClose(Cancel, 0)
    If Cancel Then Exit Sub
    Call Me.Hide
End Sub

'----------------------------------------
'◆Iniファイル
'----------------------------------------
'Iniファイルへの保存や読込の処理を記述してください
'----------------------------------------
Public Sub IniRead_UserFormInitialize()
    '------------------------------
    '◇ユーザー用Iniファイル読込処理(UserFormInitializeイベント時)
    '------------------------------
    '以下に初期化時のIniファイル読込処理を記述してください
    '------------------------------

End Sub

Public Sub IniRead_UserFormActivate()
    '------------------------------
    '◇フレームワークForm位置復帰処理
    '------------------------------
    Call Form_IniReadPosition(Me, _
        Project_IniFilePath, "Form", "Rect", False)

    '------------------------------
    '◇ユーザー用Iniファイル読込処理(UserFormActivateイベント時)
    '------------------------------
    '以下にUserForm作成初期化時のIniファイル読込処理を記述してください
    '------------------------------


End Sub

Public Sub IniWrite()
    '------------------------------
    '◇フレームワークForm位置保存処理
    '------------------------------
    Call Assert(FormProperty.Handle <> 0)

    If (FormProperty.WindowState = xlNormal) Then
        Call Form_IniWritePosition(Me, _
            Project_IniFilePath, "Form", "Rect")
    End If

    '------------------------------
    '◇ユーザー用Iniファイル書込処理
    '------------------------------
    '以下に終了時のIniファイル書込処理を記述してください
    '------------------------------


End Sub

'----------------------------------------
'◆リサイズイベント
'----------------------------------------
Private Sub UserForm_Resize()
    If FormProperty.Initializing = False Then
        '------------------------------
        '◇フレームワークアンカーレイアウト処理
        '------------------------------
        Call AnchorMenuButton.Layout

        '------------------------------
        '◇ユーザー用アンカーレイアウト処理
        '------------------------------
        '以下にユーザー独自のアンカーレイアウト処理を記述してください
        '------------------------------


    End If
End Sub

'----------------------------------------
'◆メニューボタン
'----------------------------------------
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
        PointToPixel(Me.Left + ImageMenuButton.Left + ImageMenuButton.Width) _
        + IIf(FormProperty.ResizeFrame, XOffsetResizeOn, XOffsetResizeOff) _
        - PopupMenu.Width + XOffset, _
        PointToPixel(Me.Top + ImageMenuButton.Top + ImageMenuButton.Height) _
        + IIf(FormProperty.ResizeFrame, YOffsetResizeOn, YOffsetResizeOff) _
        + IIf(FormProperty.TitleBar, YOffsetTitleBarOn, YOffsetTitleBarOff))
    Case "CreateAppShortcut"
        Call Load(FormCreateAppShortcut)
        Call FormCreateAppShortcut.ShowDialog( _
            Me, FormProperty.TopMost)
        Call Unload(FormCreateAppShortcut)
    Case "VersionInfo"
        Call MsgBox( _
            Project_VersionDialogInstruction + vbNewLine + _
            Project_VersionDialogContent, _
            vbOKOnly, _
            Project_VersionDialogWindowTitle)
    Case "AppClose"
        FormClose
    End Select
End Sub

'--------------------------------------------------
'■プログラム本体
'--------------------------------------------------
'以下にプログラム本体の処理を記述してください
'--------------------------------------------------

