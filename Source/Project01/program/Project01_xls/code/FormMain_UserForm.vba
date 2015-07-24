'--------------------------------------------------
'Excel Make App Framework
'FormMain
'
'ObjectName:    FormMain
'--------------------------------------------------
'version:       2015/03/11
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

Private AnchorMenuButton As New st_vba_ControlAnchor

'----------------------------------------
'�����[�U�[�p
'----------------------------------------
'------------------------------
'���A���J�[��`
'------------------------------
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
        '���t���[�����[�N�A���J�[����������
        '------------------------------
        Call AnchorMenuButton.Initialize( _
            Me.FrameMenuButton, _
            HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 8, 0), _
            VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 8, 0))

        '------------------------------
        '�����[�U�[�p�A���J�[����������
        '------------------------------
        '�ȉ��Ƀ��[�U�[�Ǝ��̃A���J�[�������������L�q���Ă�������
        '------------------------------

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


End Sub


'----------------------------------------
'�����T�C�Y�C�x���g
'----------------------------------------
Private Sub UserForm_Resize()
    If FormProperty.Initializing = False Then
        '------------------------------
        '���t���[�����[�N�A���J�[���C�A�E�g����
        '------------------------------
        Call AnchorMenuButton.Layout

        '------------------------------
        '�����[�U�[�p�A���J�[���C�A�E�g����
        '------------------------------
        '�ȉ��Ƀ��[�U�[�Ǝ��̃A���J�[���C�A�E�g�������L�q���Ă�������
        '------------------------------


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

    Select Case PopupMenu_PopupReturn(PopupMenu, _
        PointToPixel(Me.Left + FrameMenuButton.Left + FrameMenuButton.Width) _
        + IIf(FormProperty.ResizeFrame, 8, 4) _
        - PopupMenu.Width + 14, _
        PointToPixel(Me.Top + FrameMenuButton.Top + FrameMenuButton.Height) _
        + IIf(FormProperty.ResizeFrame, 8, 4) _
        + IIf(FormProperty.TitleBar, 20, 0))
    Case "CreateAppShortcut"
        Call Load(FormCreateAppShortcut)
        Call FormCreateAppShortcut.ShowDialog( _
            Me, FormProperty.TopMost)
        Call Unload(FormCreateAppShortcut)
    Case "VersionInfo"
        Call TaskDialogMsgBox( _
            FormProperty.Handle, _
            ExtractIcon(0, Project_MainIconFilePath, Project_MainIconIndex), _
            Project_VersionDialogWindowTitle, _
            Project_VersionDialogInstruction, _
            Project_VersionDialogContent, _
            TDCBF_OK_BUTTON, _
            True)
    Case "AppClose"
        FormClose
    End Select
End Sub

'--------------------------------------------------
'���v���O�����{��
'--------------------------------------------------
'�ȉ��Ƀv���O�����{�̂̏������L�q���Ă�������
'--------------------------------------------------
