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
        
        If IsTaskbarPinWindows Then
            CheckBoxTaskbarPin.Value = _
                fso.FileExists(Project_ShortcutFilePath_TaskbarPin)
        Else
            CheckBoxTaskbarPin.Value = False
            CheckBoxTaskbarPin.Enabled = False
        End If
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
            Project_MainIconFilePath, Project_Name, True)
            
        '[Win7AppId.exe_]を[Win7AppId.exe]に変換
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
                    'コマンド完了を待ってから実行
                    '[Win7AppId.exe]を[Win7AppId.exe_]に変換
                    Call fso.MoveFile( _
                        Project_TaskbarPinCommandExeFilePath, _
                        Project_TaskbarPinCommandExeFilePath + "_")
                End If
            End If
            
            If (CheckBoxTaskbarPin.Value = False) Then
                If FileExistsWait(Project_ShortcutFilePath_TaskbarPin, False) Then
                    'コマンド完了を待ってから実行
                    '[Win7AppId.exe]を[Win7AppId.exe_]に変換
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
'■履歴
'◇ ver 2015/07/29
'・ 作成
'◇ ver 2017/02/05
'・ Win7AppId.exe を使用する時以外は Win7AppId.exe_ とする
'--------------------------------------------------
