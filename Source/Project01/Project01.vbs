'--------------------------------------------------
'Standard Software
'Excel Make Appl Framework
'--------------------------------------------------
'�o�[�W����     2014/11/19
'--------------------------------------------------
Option Explicit
'--------------------------------------------------
'�ݒ�
Const ProgramSubFolderName = "program"
Const ProgramExcelFileName = "Project01.xls"
'--------------------------------------------------

Dim fso: Set fso  = CreateObject("Scripting.FileSystemObject")

Call Main

Sub Main
Do
    Dim ExcelApp1
    Set ExcelApp1 = CreateObject("Excel.Application")
    If ExcelApp1 Is Nothing Then
        Call MsgBox("Excel���C���X�g�[������Ă��܂���B")
        Exit Do
    End If

    ExcelApp1.IgnoreRemoteRequests = True

    ExcelApp1.Visible = False

    Call ExcelApp1.Workbooks.Open( _
        fso.GetParentFolderName(Wscript.ScriptFullName) + "\" + _
        ProgramSubFolderName + "\" + _
        ProgramExcelFileName, , True)
    '����O������ReadOnly�w��

    Dim Args: Set Args = WScript.Arguments
    Dim ArgsText: ArgsText = ""
    Dim I
    For I = 0 To Args.Count - 1
        ArgsText = ArgsText + Args(I) + vbTab
    Next

On Error Resume Next
    Call ExcelApp1.Run( _
        ProgramExcelFileName + "!Main", ArgsText)

    'Call MsgBox("End Script")

    ExcelApp1.IgnoreRemoteRequests = False
    Call ExcelApp1.Workbooks.Close()
    Call ExcelApp1.Quit
    Set ExcelApp1 = Nothing
Loop While False
End Sub

