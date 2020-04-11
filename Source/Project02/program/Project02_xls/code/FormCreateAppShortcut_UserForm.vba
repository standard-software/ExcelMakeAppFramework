'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    CreateAppShortcut Form
'ObjectName:    FormCreateAppShortcut
'--------------------------------------------------
'Version:       2020/04/11
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

    'Form位置
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

    '初期化
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
    
    'チェックボックス値設定
    With Nothing
        CheckBoxDesktop.Value = _
            fso.FileExists(Project_ShortcutFilePath_Desktop)
        CheckBoxStartMenu.Value = _
            fso.FileExists(Project_ShortcutFilePath_StartMenu)
        CheckBoxSendTo.Value = _
            fso.FileExists(Project_ShortcutFilePath_SendTo)
        
    End With
    
    'アイコン
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
            Project_MainIconFilePath, Project_Name, False)
        
    End If
End Sub

Private Sub UserForm_Activate()
    If FormProperty.Initializing Then
        FormProperty.Initializing = False
        Call FormProperty.SetWindowsProperty
    End If
End Sub
