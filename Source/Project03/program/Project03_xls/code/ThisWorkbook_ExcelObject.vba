'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Project03 ThisWorkbook
'ObjectName:    ThisWorkbook
'--------------------------------------------------
'Version:       2015/08/19
'--------------------------------------------------
Option Explicit

Public OriginalWindowRectBuffer As String

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Application.WindowState = xlMaximized Then
        Application.WindowState = xlNormal
    End If

    Call Form_IniWritePosition(Application, _
        Project_IniFilePath, _
        "Form", "Rect")

    If CanStrToRect(OriginalWindowRectBuffer) Then
        Application.Visible = False
        Call Form_SetRectPixel(Application, _
            StrToRect(OriginalWindowRectBuffer))
    End If

    Call ApplicationModeOff
    Application.DisplayAlerts = False
    Application.Quit
    Application.IgnoreRemoteRequests = False
End Sub
