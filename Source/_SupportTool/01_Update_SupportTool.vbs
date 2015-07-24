Option Explicit

'--------------------------------------------------
'��Include st.vbs
'--------------------------------------------------
Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.GetParentFolderName(WScript.ScriptFullName) _
        + "\" + FileName, 1)
    Call ExecuteGlobal(Stream.ReadAll())
    Call Stream.Close
End Sub
'--------------------------------------------------
Call Include(".\Lib\st.vbs")
'--------------------------------------------------

'------------------------------
'�����C������
'------------------------------
Call Main

Sub Main
    Dim MessageText: MessageText = ""

    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(ScriptFolderPath, "SupportTool.ini"))

    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)

    '--------------------
    '�E�ݒ�Ǎ�
	'--------------------
    Dim SupportTool_Source_Path: SupportTool_Source_Path = _
        IniFile.ReadString("Update_SupportTool", "SupportToolSourcePath", "")
    If SupportTool_Source_Path = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    Dim SupportTool_IgnoreFile: SupportTool_IgnoreFile = _
        IniFile.ReadString("Update_SupportTool", "SupportToolIgnoreFiles", "")
    '--------------------

    Dim SourceFolderPath: SourceFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, SupportTool_Source_Path)
    If not fso.FolderExists(SourceFolderPath) Then
        WScript.Echo _
            "�R�s�[���t�H���_��������܂���" + vbCrLF + _
            SourceFolderPath
        Exit Sub
    End If

    Dim DestFolderPath: DestFolderPath = _
        ScriptFolderPath

    If LCase(SourceFolderPath) = LCase(DestFolderPath) Then
        WScript.Echo _
            "�R�s�[��ƃR�s�[���̃t�H���_������ł��B" + vbCrLF + _
            SourceFolderPath
        Exit Sub
    End If

'    Call CopyFolderOverWriteIgnore( _
'        SourceFolderPath, DestFolderPath, "*.ini")

    Call DeleteFileTargetPath( _
        DestFolderPath, "*.vbs")

    Call CopyFolderIgnorePath( _
        SourceFolderPath, DestFolderPath, _
        StringCombine(",", Array("*.ini", "Update_HereLib.vbs", SupportTool_IgnoreFile)), _
        "")

    MessageText = MessageText + _
        DestFolderPath + vbCrLf

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)
End Sub

