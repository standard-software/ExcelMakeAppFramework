Option Explicit

'--------------------------------------------------
'■Include st.vbs
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
'◇メイン処理
'------------------------------
Call Main

Sub Main
    Dim MessageText: MessageText = ""

    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(ScriptFolderPath, "SupportTool.ini"))

    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)

    '------------------------------
    '・設定読込
    '------------------------------
    Dim BackupSourceFolderPaths: BackupSourceFolderPaths = _
        IniFile.ReadString("SourceBackup", "BackupSourceFolderPaths", "")
    If BackupSourceFolderPaths = "" Then
        WScript.Echo _
            "設定が読み取れていません"
        Exit Sub
    End If

    Dim BackupDestFolderPaths: BackupDestFolderPaths = _
        IniFile.ReadString("SourceBackup", "BackupDestFolderPaths", "")
    If BackupDestFolderPaths = "" Then
        WScript.Echo _
            "設定が読み取れていません"
        Exit Sub
    End If
    'BackupSourceFolderPaths と BackupDestFolderPaths は
    'カンマ区切りで同じ個数のフォルダ指定とする

    Dim BackupFolderLastYYYY_MM_DD: BackupFolderLastYYYY_MM_DD = _
        IniFile.ReadString("SourceBackup", "BackupFolderLastYYYY_MM_DD", "True")
    If LCase(BackupFolderLastYYYY_MM_DD) = "true" Then
        BackupFolderLastYYYY_MM_DD = True
    Else
        BackupFolderLastYYYY_MM_DD = False
    End If
    '------------------------------
    Dim BackupSourceFolderPathArray
    BackupSourceFolderPathArray = Split(BackupSourceFolderPaths, ",")
    Dim BackupDestFolderPathArray
    BackupDestFolderPathArray = Split(BackupDestFolderPaths, ",")

    If ArrayCount(BackupSourceFolderPathArray) _
    <> ArrayCount(BackupDestFolderPathArray) Then
        WScript.Echo _
            "BackupSourceFolderPaths と BackupDestFolderPaths の" + vbCrLf + _
            "指定個数が正しくありません。" + vbCrLf + _
            "処理を中止します。"
        Exit Sub
    End If

    Dim NowValue: NowValue = Now

    Dim BackupFolderPath
    Dim SourceFolderPath

    Dim I
    For I = 0 To ArrayCount(BackupSourceFolderPathArray) - 1
    Do
        SourceFolderPath = _
            BackupSourceFolderPathArray(I)
        SourceFolderPath = _
            AbsoluteFilePath(ScriptFolderPath, SourceFolderPath)

        If not fso.FolderExists(SourceFolderPath) Then
            WScript.Echo _
                "Sourceフォルダが見つかりません。" + vbCrLF + _
                SourceFolderPath
            Exit Do
        End If

        BackupFolderPath = BackupDestFolderPathArray(I)
        BackupFolderPath = _
            AbsoluteFilePath(ScriptFolderPath, BackupFolderPath)

        If BackupFolderLastYYYY_MM_DD Then
            '[Backup\Source\2015-02-27]の形式
            BackupFolderPath = _
                PathCombine(Array( _
                    BackupFolderPath, _
                    FormatYYYY_MM_DD(NowValue, "-") + _
                    "_" + _
                    FormatHH_MM_SS(NowValue, "-")))
        Else
            '[Backup\2015-02-27\Source]の形式
            BackupFolderPath = _
                PathCombine(Array( _
                    fso.GetParentFolderName(BackupFolderPath), _
                    FormatYYYY_MM_DD(NowValue, "-") + _
                    "_" + _
                    FormatHH_MM_SS(NowValue, "-"), _
                    fso.GetFileName(BackupFolderPath)))
        End If

        Call ForceCreateFolder(BackupFolderPath)

        Dim Folders: Folders = Split( _
            FolderPathListTopFolder(SourceFolderPath), vbCrLf)
        Dim Folder
        For Each Folder In Folders
            Call fso.CopyFolder(Folder, _
                PathCombine(Array(BackupFolderPath, _
                fso.GetFileName(Folder))), True)
            If BackupFolderLastYYYY_MM_DD Then
                MessageText = MessageText + fso.GetFileName(Folder) + vbCrLf
            End If
        Next

        Dim Files: Files = Split( _
            FilePathListTopFolder(SourceFolderPath), vbCrLf)
        Dim File
        For Each File In Files
            Call fso.CopyFile(File, _
                PathCombine(Array(BackupFolderPath, _
                fso.GetFileName(File))), True)
        Next
        If BackupFolderLastYYYY_MM_DD = False Then
            MessageText = MessageText + fso.GetFileName(SourceFolderPath) + vbCrLf
        End If

    Loop While False
    Next

    WScript.Echo _
        "Finish " + WScript.ScriptName + vbCrLf + _
        "----------" + vbCrLf + _
        Trim(MessageText)

End Sub

