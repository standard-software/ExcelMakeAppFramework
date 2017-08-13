'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Main Form
'ObjectName:    FormMain
'--------------------------------------------------
'Version:       2015/07/29
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'���錾
'--------------------------------------------------
'----------------------------------------
'���t���[�����[�N�p
'----------------------------------------
Public Args As String

Private FormProperty As New st_vba_FormProperty

Private FAnchorMenuButton As New st_vba_ControlAnchor

'----------------------------------------
'�����[�U�[�p
'----------------------------------------
'------------------------------
'���A���J�[��`
'------------------------------
Private FAnchorLeftTextBox As New st_vba_ControlAnchor
Private FAnchorTopTextBox As New st_vba_ControlAnchor
Private FAnchorBottomTextBox As New st_vba_ControlAnchor
Private FAnchorSplitter1 As New st_vba_ControlAnchor
Private FAnchorSplitter2 As New st_vba_ControlAnchor

Private FSplitter1 As New st_vba_ControlSplitter
Private FSplitter2 As New st_vba_ControlSplitter
'------------------------------
'���ϐ���`
'------------------------------

'--------------------------------------------------
'������
'--------------------------------------------------

'----------------------------------------
'���N���E�I��
'----------------------------------------

'------------------------------
'���ϐ��������Ȃ�
'------------------------------
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    Args = ""
    Call IniRead_UserFormInitialize
End Sub

'------------------------------
'��Main����̌Ăяo��
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
    '���t���[�����[�N����������
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
    '�����[�U�[�p����������
    '------------------------------
    '�ȉ��Ƀ��[�U�[�Ǝ��̏������������L�q���Ă�������
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
        '�����j���[�{�^�����E��[�ɂ���
        '------------------------------
        Me.FrameMenuButton.Top = 0
        Me.FrameMenuButton.Left = _
            Me.FrameMenuButton.Parent.InsideWidth - _
            Me.FrameMenuButton.Width + 1

        '------------------------------
        '���t���[�����[�N�A���J�[����������
        '------------------------------
        Call FAnchorMenuButton.Initialize( _
            Me.FrameMenuButton, _
            HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 0, 0))
        'Excel2016�ł́AOffset�l��ResizeFrame�ɂ�����炸0�ɂȂ�
        'Excel2013�ł͉��L�̃R�[�h���L��
        'Call FAnchorMenuButton.Initialize( _
        '   Me.FrameMenuButton, _
        '   HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 8, 0), _
        '   VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 8, 0))

        '------------------------------
        '�����[�U�[�p�A���J�[����������
        '------------------------------
        '�ȉ��Ƀ��[�U�[�Ǝ��̃A���J�[�������������L�q���Ă�������
        '------------------------------
        Call FAnchorLeftTextBox.Initialize( _
            Me.TextBoxLeft, _
            HorizonAnchorType.haLeft, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaStretch, IIf(FormProperty.ResizeFrame, 0, 0))
        Call FAnchorTopTextBox.Initialize( _
            Me.TextBoxTop, _
            HorizonAnchorType.haStretch, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 0, 0))
        Call FAnchorBottomTextBox.Initialize( _
            Me.TextBoxBottom, _
            HorizonAnchorType.haStretch, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaStretch, IIf(FormProperty.ResizeFrame, 0, 0))
        Call FAnchorSplitter1.Initialize( _
            Me.ImageSplitter1, _
            HorizonAnchorType.haLeft, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaStretch, IIf(FormProperty.ResizeFrame, 0, 0))
        Call FAnchorSplitter2.Initialize( _
            Me.ImageSplitter2, _
            HorizonAnchorType.haStretch, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 0, 0))

        Call FSplitter1.Initialize( _
            ImageSplitter1, _
            SplitterType.Vertical, _
            10, 10)
        Call FSplitter1.AddControlLeftTop(TextBoxLeft)
        Call FSplitter1.AddControlRightBottom(TextBoxTop)
        Call FSplitter1.AddControlRightBottom(TextBoxBottom)
        Call FSplitter1.AddControlRightBottom(ImageSplitter2)
        
        Call FSplitter2.Initialize( _
            ImageSplitter2, _
            SplitterType.Horizon, _
            10, 10)
        Call FSplitter2.AddControlLeftTop(TextBoxTop)
        Call FSplitter2.AddControlRightBottom(TextBoxBottom)

        Call IniRead_UserFormActivate

        '���C�A�E�g�A���J�[�𓮍삳����
        Call UserForm_Resize

        Call FormProperty.ForceActiveMouseClick

        '------------------------------
        '����������
        '------------------------------
        '�����ɑ΂��鏈�����L�q���Ă�������
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
'���I�����ɌĂяo���֐�
'------------------------------
'Me.Hide �� Call Unload(Me) �ł͂Ȃ�
'����FormClose�֐����Ăяo���Ă�������
'Me.Hide �� Call Unload(Me) �ł�
'UserForm_QueryClose�C�x���g���Ăяo���ꂸ
'Ini�t�@�C���ւ̕ۑ����s���܂���B
'------------------------------
Private Sub FormClose()
    Dim Cancel As Integer
    Cancel = False
    Call UserForm_QueryClose(Cancel, 0)
    If Cancel Then Exit Sub
    Call Me.Hide
End Sub

'----------------------------------------
'��Ini�t�@�C��
'----------------------------------------
'Ini�t�@�C���ւ̕ۑ���Ǎ��̏������L�q���Ă�������
'----------------------------------------
Public Sub IniRead_UserFormInitialize()
    '------------------------------
    '�����[�U�[�pIni�t�@�C���Ǎ�����(UserFormInitialize�C�x���g��)
    '------------------------------
    '�ȉ��ɏ���������Ini�t�@�C���Ǎ��������L�q���Ă�������
    '------------------------------

End Sub

Public Sub IniRead_UserFormActivate()
    '------------------------------
    '���t���[�����[�NForm�ʒu���A����
    '------------------------------
    Call Form_IniReadPosition(Me, _
        Project_IniFilePath, "Form", "Rect", False)

    '------------------------------
    '�����[�U�[�pIni�t�@�C���Ǎ�����(UserFormActivate�C�x���g��)
    '------------------------------
    '�ȉ���UserForm�쐬����������Ini�t�@�C���Ǎ��������L�q���Ă�������
    '------------------------------
    Dim Splitter1Left As Long
    Splitter1Left = StrToLongDefault( _
        IniFile_GetString(Project_IniFilePath, _
            "Form", "Splitter1Left"), ImageSplitter1.Left)
            
    If FSplitter1.CanLayoutUpdate(Splitter1Left, ImageSplitter1.Top) Then
        Call FSplitter1.LayoutUpdate(Splitter1Left, ImageSplitter1.Top)
    End If
    
    Dim Splitter2Top As Long
    Splitter2Top = StrToLongDefault( _
        IniFile_GetString(Project_IniFilePath, _
            "Form", "Splitter2Top"), ImageSplitter2.Top)
            
    If FSplitter2.CanLayoutUpdate(ImageSplitter2.Left, Splitter2Top) Then
        Call FSplitter2.LayoutUpdate(ImageSplitter2.Left, Splitter2Top)
    End If
End Sub

Public Sub IniWrite()
    '------------------------------
    '���t���[�����[�NForm�ʒu�ۑ�����
    '------------------------------
    Call Assert(FormProperty.Handle <> 0)

    If (FormProperty.WindowState = xlNormal) Then
        Call Form_IniWritePosition(Me, _
            Project_IniFilePath, "Form", "Rect")
    End If

    '------------------------------
    '�����[�U�[�pIni�t�@�C����������
    '------------------------------
    '�ȉ��ɏI������Ini�t�@�C�������������L�q���Ă�������
    '------------------------------
    If (FormProperty.WindowState = xlNormal) Then
        Call IniFile_SetString(Project_IniFilePath, _
            "Form", "Splitter1Left", _
            ImageSplitter1.Left)
        Call IniFile_SetString(Project_IniFilePath, _
            "Form", "Splitter2Top", _
            ImageSplitter2.Top)
    End If
End Sub

'----------------------------------------
'�����T�C�Y�C�x���g
'----------------------------------------
Private Sub UserForm_Resize()
    If FormProperty.Initializing = False Then
        '------------------------------
        '���t���[�����[�N�A���J�[���C�A�E�g����
        '------------------------------
        Call FAnchorMenuButton.Layout

        '------------------------------
        '�����[�U�[�p�A���J�[���C�A�E�g����
        '------------------------------
        '�ȉ��Ƀ��[�U�[�Ǝ��̃A���J�[���C�A�E�g�������L�q���Ă�������
        '------------------------------
        Call FAnchorLeftTextBox.Layout
        Call FAnchorTopTextBox.Layout
        Call FAnchorBottomTextBox.Layout
        Call FAnchorSplitter1.Layout
        Call FAnchorSplitter2.Layout
    End If
End Sub

'----------------------------------------
'�����j���[�{�^��
'----------------------------------------
Private Sub ImageMenuButton_Click()
    Dim PopupMenu As CommandBar
    Set PopupMenu = Application.CommandBars.Add(, Position:=msoBarPopup)

    Dim MenuItemCreateAppShortcut As CommandBarControl
    Set MenuItemCreateAppShortcut = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemCreateAppShortcut.Caption = "�A�v���P�[�V�����̃V���[�g�J�b�g���쐬..."
    MenuItemCreateAppShortcut.FaceId = 0
    MenuItemCreateAppShortcut.OnAction = PopupMenu_ActionText("CreateAppShortcut")

    Dim MenuItemVersionInfo As CommandBarControl
    Set MenuItemVersionInfo = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemVersionInfo.Caption = "�o�[�W�������"
    MenuItemVersionInfo.FaceId = 0
    MenuItemVersionInfo.OnAction = PopupMenu_ActionText("VersionInfo")

    Dim MenuItemAppClose As CommandBarControl
    Set MenuItemAppClose = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemAppClose.BeginGroup = True
    MenuItemAppClose.Caption = "�I��"
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
        PointToPixel(Me.Left + FrameMenuButton.Left + FrameMenuButton.Width) _
        + IIf(FormProperty.ResizeFrame, XOffsetResizeOn, XOffsetResizeOff) _
        - PopupMenu.Width + XOffset, _
        PointToPixel(Me.Top + FrameMenuButton.Top + FrameMenuButton.Height) _
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
'���v���O�����{��
'--------------------------------------------------
'�ȉ��Ƀv���O�����{�̂̏������L�q���Ă�������
'--------------------------------------------------

