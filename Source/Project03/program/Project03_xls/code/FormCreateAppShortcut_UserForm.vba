'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    CreateAppShortcut Form
'ObjectName:    FormCreateAppShortcut
'--------------------------------------------------
'Version:       2017/02/05
'--------------------------------------------------
Option Explicit

Public ModalResult As VbMsgBoxResult

Public FormProperty As New st_vba_FormProperty

Public ParentForm As Object

Private Sub ButtonOK_Click()
    ModalResult = vbOK
    Me.Hide
End Sub

Private Sub ButtonCancel_Click()
    ModalResult = vbCancel
    Me.Hide
End Sub

Sub ShowDialog( _
Optional ByVal ParentForm As Object = Nothing, _
Optional ByVal TopMost As Boolean = False)

    'Form�ʒu
    If (ParentForm Is Nothing) = False Then
        Me.StartUpPosition = 0
        Dim FormRect As Rect
        FormRect = Form_GetRectPixel(Me)
        FormRect = GetRectMoveCenter(FormRect, _
                GetPointRectCenter(Form_GetRectPixel(ParentForm)))
        Call Form_SetRectPixel(Me, GetRectInsideDesktopRect(FormRect, GetRectWorkArea))
    Else
        Me.StartUpPosition = 1
    End If

    '������
    Call FormProperty.InitializeForm(Me)
    
    Call FormProperty.InitializeProperty( _
        TaskBarButton:=False, _
        TitleBar:=True, _
        SystemMenu:=True, _
        FormIcon:=False, _
        MinimizeButton:=False, _
        MaximizeButton:=False, _
        CloseButton:=True, _
        ResizeFrame:=False, _
        TopMost:=True)
    Me.Caption = Project_FormCreateAppShortcut_Title
    
    '�`�F�b�N�{�b�N�X�l�ݒ�
    With Nothing
        CheckBoxDesktop.Value = _
            fso.FileExists(Project_ShortcutFilePath_Desktop)
        CheckBoxStartMenu.Value = _
            fso.FileExists(Project_ShortcutFilePath_StartMenu)
        CheckBoxSendTo.Value = _
            fso.FileExists(Project_ShortcutFilePath_SendTo)
        
        If IsTaskbarPinWindows Then
            CheckBoxTaskbarPin.Value = _
                fso.FileExists(Project_ShortcutFilePath_TaskbarPin)
        Else
            CheckBoxTaskbarPin.Value = False
            CheckBoxTaskbarPin.Enabled = False
        End If
    End With
    
    '�A�C�R��
    Dim hBitmap As Long
    hBitmap = GetBitmapDrawIcon( _
        NewIconFilePathIndex(SystemIconFilePath, ID_ICON_INFORMATION), _
        NewRectSize(32, 32))
    Call Image_Picture_SetBitmap(Image1, hBitmap)
    
    
    Call Me.Show(vbModal)
    
    If ModalResult = vbOK Then
        
        Call SetShortcutIcon(CheckBoxDesktop.Value, _
            Project_ShortcutFilePath_Desktop, _
            ProjectScriptFilePath, _
            Project_MainIconFilePath, Project_Name, False)
        
        Call SetShortcutIcon(CheckBoxStartMenu.Value, _
            Project_ShortcutFilePath_StartMenu, _
            ProjectScriptFilePath, _
            Project_MainIconFilePath, Project_Name, True)
            
        Call SetShortcutIcon(CheckBoxSendTo.Value, _
            Project_ShortcutFilePath_SendTo, _
            ProjectScriptFilePath, _
            Project_MainIconFilePath, Project_Name, True)
            
        '[Win7AppId.exe_]��[Win7AppId.exe]�ɕϊ�
        If fso.FileExists(Project_TaskbarPinCommandExeFilePath) = False Then
            If fso.FileExists(Project_TaskbarPinCommandExeFilePath + "_") Then
                Call fso.MoveFile( _
                    Project_TaskbarPinCommandExeFilePath + "_", _
                    Project_TaskbarPinCommandExeFilePath)
            Else
                Call Assert(False, "Error:Win7AppId.exe_ is not exist.")
            End If
        End If
            
        If FileExistsWait(Project_TaskbarPinCommandExeFilePath) Then
            Call SetTaskbarPinShortcutIcon(CheckBoxTaskbarPin.Value, _
                Project_ShortcutFilePath_TaskbarPin, _
                ProjectScriptFilePath, _
                Project_MainIconFilePath, Project_Name, _
                PathCombine(GetSpecialFolderPath(System), "cscript.exe"), _
                "Microsoft " + ChrW(&HAE) + " Console Based Script Host.lnk", _
                Project_TaskbarPinCommandExeFilePath, _
                Project_AppID)
            
            If (CheckBoxTaskbarPin.Value) Then
                If FileExistsWait(Project_ShortcutFilePath_TaskbarPin) Then
                    '�R�}���h������҂��Ă�����s
                    '[Win7AppId.exe]��[Win7AppId.exe_]�ɕϊ�
                    Call fso.MoveFile( _
                        Project_TaskbarPinCommandExeFilePath, _
                        Project_TaskbarPinCommandExeFilePath + "_")
                End If
            End If
            
            If (CheckBoxTaskbarPin.Value = False) Then
                If FileExistsWait(Project_ShortcutFilePath_TaskbarPin, False) Then
                    '�R�}���h������҂��Ă�����s
                    '[Win7AppId.exe]��[Win7AppId.exe_]�ɕϊ�
                    Call fso.MoveFile( _
                        Project_TaskbarPinCommandExeFilePath, _
                        Project_TaskbarPinCommandExeFilePath + "_")
                End If
            End If
            
        End If
        
    End If
End Sub

Private Sub UserForm_Activate()
    If FormProperty.Initializing Then
        FormProperty.Initializing = False
        Call FormProperty.SetWindowsProperty
    End If
End Sub

'--------------------------------------------------
'������
'�� ver 2015/07/29
'�E �쐬
'�� ver 2017/02/05
'�E Win7AppId.exe ���g�p���鎞�ȊO�� Win7AppId.exe_ �Ƃ���
'--------------------------------------------------
