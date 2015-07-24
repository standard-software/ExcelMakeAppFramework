'--------------------------------------------------
'st.vbs
'--------------------------------------------------
'ModuleName:    Basic_Module
'FileName:      st.vbs
'--------------------------------------------------
'Discription:   Standard Software Library For VBScript
'--------------------------------------------------
'OpenSource:    https://github.com/standard-software/st.vbs
'License:       MIT License
'   URL:        https://github.com/standard-software/st.vbs/blob/master/Document/Readme.txt
'All Right Reserved:
'   Name:       Standard Software
'   URL:        http://standard-software.net/
'--------------------------------------------------
'Version:       2015/07/24
'--------------------------------------------------

'--------------------------------------------------
'■マーク
'--------------------------------------------------

    '--------------------------------------------------
    '■
    '--------------------------------------------------

    '----------------------------------------
    '◆
    '----------------------------------------
    '◇
    '----------------------------------------
    '・
    '----------------------------------------

Option Explicit

'--------------------------------------------------
'■定数/型宣言
'--------------------------------------------------

'----------------------------------------
'◆FileSystemObject
'----------------------------------------
Public fso: Set fso = CreateObject("Scripting.FileSystemObject")

'----------------------------------------
'◆Shell
'----------------------------------------
Public Shell: Set Shell = WScript.CreateObject("WScript.Shell")

'--------------------------------------------------
'■実装
'--------------------------------------------------

'----------------------------------------
'◆テスト
'----------------------------------------
'Call test
Public Sub test
'    Call WScript.Echo("test")
'    Call WScript.Echo(vbObjectError)
''    Call testAssert
'    Call testCheck
'    Call testOrValue
    Call testSetValue
'    Call testLoadTextFile
'    Call testSaveTextFile
'    Call testFirstStrFirstDelim
'    Call testFirstStrLastDelim
'    Call testLastStrFirstDelim
'    Call testLastStrLastDelim
'    Call testIsFirstStr()
'    Call testIncludeFirstStr()
'    Call testExcludeFirstStr()
'    Call testIsLastStr()
'    Call testIncludeLastStr()
'    Call testExcludeLastStr()
'    Call testTrimFirstStrs
'    Call testTrimLastStrs
    Call testStringCombine

'    Call testAbsoluteFilePath
'    Call testPeriodExtName
'    Call testExcludePathExt
'    Call testChangeFileExt
'    Call testPathCombine

'    Call testFileFolderPathList
    'テストフォルダが必要なので注意

'    Call testShellCommandRunReturn
'    Call testEnvironmentalVariables
'    Call testShellFileOpen
'    Call testShellCommandRun

'    Call testFormatYYYYMMDD
'    Call testFormatHHMMSS
    Call testFormatDateTimeToString

'    Call testIniFile
    Call testMatchText
    Call testMatchTextWildCard
    Call testMatchTextKeyWord
    Call testMatchTextRegExp

    Call testDeleteDateText
    Call testDeleteDateTimeText

'    Call testForceCreateFolder
    'テストフォルダを作成してしまうので注意

    Call testMaxValue
    Call testMinValue
    Call testIsLong
    Call testLongToStrDigitZero
    Call testStrToLongDefault

    Call testArrayCount
    Call testArrayAdd
    Call testArrayDelete
    Call testArrayInsert
    Call testArrayFunctions

    Call testProcessExists

    Call testSplit

    Call testArgsOption
    Call testArgsOptionCheckError

Call MsgBox("st.vbs Test Finish")
End Sub

'----------------------------------------
'◆条件判断
'----------------------------------------

'----------------------------------------
'・Assert
'----------------------------------------
'   ・  Assert = 主張する
'   ・  Err番号 vbObjectError + 1 は
'       ユーザー定義エラー番号の1
'       動作させると次のようにエラーダイアログが表示される
'       エラー: Messageの内容
'       コード: 80040001
'       ソース: Sub Assert
'----------------------------------------
Public Sub Assert(ByVal Value, ByVal Message)
    If Value = False Then
        Call Err.Raise(vbObjectError + 1, "Sub Assert", Message)
    End If
End Sub

Private Sub testAssert()
    Call Assert(False, "TEST")
End Sub

'----------------------------------------
'・Check
'----------------------------------------
'   ・  2つの値を比較して一致しなければ
'       メッセージを出す関数
'----------------------------------------
Public Function Check(ByVal A, ByVal B)
    Check = (A = B)
    If Check = False Then
        Call WScript.Echo("A <> B" + vbCrLf + _
            "A = " + CStr(A) + vbCrLf + _
            "B = " + CStr(B))
    End If
End Function

Private Sub testCheck()
    Call Check(10, 20)
End Sub

'----------------------------------------
'・OrValue
'----------------------------------------
'   ・例：If OrValue(ValueA, Array(1, 2, 3)) Then
'----------------------------------------
Public Function OrValue(ByVal Value, ByVal Values)
    Call Assert(IsArray(Values), "Error:OrValue:Values is not Array.")

    OrValue = False
    Dim I
    For I = LBound(Values) To UBound(Values)
        If Value = Values(I) Then
            OrValue = True
            Exit For
        End If
    Next
End Function

Private Sub testOrValue()
    Call Check(True, OrValue(10, Array(20, 30, 40, 10)))
    Call Check(False, OrValue(50, Array(20, 30, 40, 10)))
End Sub

'----------------------------------------
'・IIF
'----------------------------------------
Public Function IIF(ByVal CompareValue, ByVal Result1, ByVal Result2)
    If CompareValue Then
        IIF = Result1
    Else
        IIf = Result2
    End If
End Function

'----------------------------------------
'◆型、型変換
'----------------------------------------

'----------------------------------------
'・変数に値やオブジェクトをセットする
'----------------------------------------
Public Sub SetValue(ByRef Variable, ByVal Value)
    If IsObject(Value) Then
        Set Variable = Value
    Else
        Variable = Value
    End If
End Sub

Private Sub testSetValue
    Dim A
    A = 1
    Call SetValue(A, 2)
    Call Check(2, A)

    Dim B
    Call SetValue(B, fso)
    Call Check("test.txt", B.GetFileName("C:\temp\test.txt"))
End Sub

'----------------------------------------
'・Long
'----------------------------------------
Public Function IsLong(ByVal Value)
    Dim Result: Result = False
    If IsNumeric(Value) Then
        If CInt(Value) = CDbl(Value) Then
            Result = True
        End If
    End If
    IsLong = Result
End Function

Private Sub testIsLong()
    Call Check(True, IsLong("123"))
    Call Check(False, IsLong("12a"))
    Call Check(False, IsLong("123.4"))
End Sub

Public Function LongToStrDigitZero(ByVal Value, ByVal Digit)
    Dim Result: Result = ""
    If 0 <= Value Then
        Result = String(MaxValue(Array(0, Digit - Len(CStr(Value)))), "0") + CStr(Value)
    Else
        Result = "-" + String(Digit - Len(CStr(Abs(Value))), "0") + CStr(Abs(Value))
    End If
    LongToStrDigitZero = Result
End Function

Private Sub testLongToStrDigitZero()
    Call Check("003", LongToStrDigitZero(3, 3))
    Call Check("000", LongToStrDigitZero(0, 3))
    Call Check("1000", LongToStrDigitZero(1000, 3))
    Call Check("-050", LongToStrDigitZero(-50, 3))
End Sub

Public Function StrToLongDefault(ByVal S, ByVal Default)
On Error Resume Next
    Dim Result
    Result = Default
    Result = CLng(S)
    StrToLongDefault = Result
End Function

Private Sub testStrToLongDefault()
    Call Check(123, StrToLongDefault("123", 0))
    Call Check(123, StrToLongDefault(" 123 ", 0))
    Call Check(0, StrToLongDefault(" A123 ", 0))
    Call Check(123, StrToLongDefault("BBB", 123))
End Sub

'----------------------------------------
'◇Boolean
'----------------------------------------
Function StrToBoolean(ByVal Value, ByVal DefaultValue)
    Call Assert(OrValue(DefaultValue, Array(True, False)), "Error:StrToBoolean")

    Dim Result: Result = DefaultValue

    If LCase(Value) = "true" Then
        Result = True
    ElseIf LCase(Value) = "false" Then
        Result = False
    Else
        Result = DefaultValue
    End If

    StrToBoolean = Result
End Function

Sub CheckStrToBoolean
    Call Check(True, StrToBoolean("True", False))
    Call Check(True, StrToBoolean("true", False))
    Call Check(True, StrToBoolean("TRUE", False))
    Call Check(False, StrToBoolean("False", True))
    Call Check(False, StrToBoolean("false", True))
    Call Check(False, StrToBoolean("FALSE", True))

    Call Check(True, StrToBoolean("", True))
    Call Check(True, StrToBoolean("123", True))
    Call Check(True, StrToBoolean("ABC", True))
    Call Check(False, StrToBoolean("", False))
    Call Check(False, StrToBoolean("123", False))
    Call Check(False, StrToBoolean("ABC", False))

End Sub

'----------------------------------------
'◆数値処理
'----------------------------------------

'----------------------------------------
'・最大値最小値
'----------------------------------------
'   ・  例：MsgBox MaxValue(Array(1, 2, 3))
'----------------------------------------
Public Function MaxValue(ByVal Values)
    Call Assert(IsArray(Values), "Error:OrValue:Values is not Array.")
    Dim Result: Result = Empty
    Dim Value
    For Each Value In Values
        If IsEmpty(Result) Then
            Result = Value
        ElseIf Result < Value Then
            Result = Value
        End If
    Next
    MaxValue = Result
End Function

Private Sub testMaxValue()
    Call Check(100, MaxValue(Array(50, 20, 30, 100, 9)))
End Sub

Public Function MinValue(ByVal Values)
    Dim Result: Result = Empty
    Dim Value
    For Each Value In Values
        If Result = Empty Then
            Result = Value
        ElseIf Result > Value Then
            Result = Value
        End If
    Next
    MinValue = Result
End Function

Private Sub testMinValue()
    Call Check(9, MinValue(Array(50, 20, 30, 100, 9)))
End Sub

'----------------------------------------
'・値範囲
'----------------------------------------
Public Function InRange(ByVal MinValue, ByVal Value, ByVal MaxValue)
    InRange = ((MinValue <= Value) And (Value <= MaxValue))
End Function

'----------------------------------------
'◆文字列処理
'----------------------------------------

'----------------------------------------
'・IsInclude
'----------------------------------------
Public Function IsIncludeStr(ByVal Str, ByVal SubStr)
    IsIncludeStr = _
        (1 <= InStr(Str, SubStr))
End Function

'----------------------------------------
'◇First / Last
'----------------------------------------

'----------------------------------------
'・FirstStr
'----------------------------------------
Public Function IsFirstStr(ByVal Str , ByVal SubStr)
    Dim Result: Result = False
    Do
        If SubStr = "" Then Exit Do
        If Str = "" Then Exit Do
        If Len(Str) < Len(SubStr) Then Exit Do

        If InStr(1, Str, SubStr) = 1 Then
            Result = True
        End If
    Loop While False
    IsFirstStr = Result
End Function

Private Sub testIsFirstStr()
    Call Check(True, IsFirstStr("12345", "1"))
    Call Check(True, IsFirstStr("12345", "12"))
    Call Check(True, IsFirstStr("12345", "123"))
    Call Check(False, IsFirstStr("12345", "23"))
    Call Check(False, IsFirstStr("", "34"))
    Call Check(False, IsFirstStr("12345", ""))
    Call Check(False, IsFirstStr("123", "1234"))
End Sub

Public Function IncludeFirstStr(ByVal Str, ByVal SubStr)
    If IsFirstStr(Str, SubStr) Then
        IncludeFirstStr = Str
    Else
        IncludeFirstStr = SubStr + Str
    End If
End Function

Private Sub testIncludeFirstStr()
    Call Check("12345", IncludeFirstStr("12345", "1"))
    Call Check("12345", IncludeFirstStr("12345", "12"))
    Call Check("12345", IncludeFirstStr("12345", "123"))
    Call Check("2312345", IncludeFirstStr("12345", "23"))
End Sub

Public Function ExcludeFirstStr(ByVal Str, ByVal SubStr)
    If IsFirstStr(Str, SubStr) Then
        ExcludeFirstStr = Mid(Str, Len(SubStr) + 1)
    Else
        ExcludeFirstStr = Str
    End If
End Function

Private Sub testExcludeFirstStr()
    Call Check("2345", ExcludeFirstStr("12345", "1"))
    Call Check("345", ExcludeFirstStr("12345", "12"))
    Call Check("45", ExcludeFirstStr("12345", "123"))
    Call Check("12345", ExcludeFirstStr("12345", "23"))
End Sub

'----------------------------------------
'・FirstText
'----------------------------------------
'   ・  大小文字を区別せずに比較する
'----------------------------------------
Public Function IsFirstText(ByVal Str , ByVal SubStr)
    IsFirstText = IsFirstStr(LCase(Str), LCase(SubStr))
End Function

Public Function IncludeFirstText(ByVal Str , ByVal SubStr)
    If IsFirstText(Str, SubStr) Then
        IncludeFirstText = Str
    Else
        IncludeFirstText = SubStr + Str
    End If
End Function

Public Function ExcludeFirstText(ByVal Str, ByVal SubStr)
    If IsFirstText(Str, SubStr) Then
        ExcludeFirstText = Mid(Str, Len(SubStr) + 1)
    Else
        ExcludeFirstText = Str
    End If
End Function

'----------------------------------------
'・LastStr
'----------------------------------------
Public Function IsLastStr(ByVal Str, ByVal SubStr)
    Dim Result: Result = False
    Do
        If SubStr = "" Then Exit Do
        If Str = "" Then Exit Do
        If Len(Str) < Len(SubStr) Then Exit Do

        If Right(Str, Len(SubStr)) = SubStr Then
            Result = True
        End If
    Loop While False
    IsLastStr = Result
End Function

Private Sub testIsLastStr()
    Call Check(True, IsLastStr("12345", "5"))
    Call Check(True, IsLastStr("12345", "45"))
    Call Check(True, IsLastStr("12345", "345"))
    Call Check(False, IsLastStr("12345", "34"))
    Call Check(False, IsLastStr("", "34"))
    Call Check(False, IsLastStr("12345", ""))
    Call Check(False, IsLastStr("123", "1234"))
End Sub

Public Function IncludeLastStr(ByVal Str, ByVal SubStr)
    If IsLastStr(Str, SubStr) Then
        IncludeLastStr = Str
    Else
        IncludeLastStr = Str + SubStr
    End If
End Function

Private Sub testIncludeLastStr()
    Call Check("12345", IncludeLastStr("12345", "5"))
    Call Check("12345", IncludeLastStr("12345", "45"))
    Call Check("12345", IncludeLastStr("12345", "345"))
    Call Check("1234534", IncludeLastStr("12345", "34"))
End Sub

Public Function ExcludeLastStr(ByVal Str, ByVal SubStr)
    If IsLastStr(Str, SubStr) Then
        ExcludeLastStr = Mid(Str, 1, Len(Str) - Len(SubStr))
    Else
        ExcludeLastStr = Str
    End If
End Function

Private Sub testExcludeLastStr()
    Call Check("1234", ExcludeLastStr("12345", "5"))
    Call Check("123", ExcludeLastStr("12345", "45"))
    Call Check("12", ExcludeLastStr("12345", "345"))
    Call Check("12345", ExcludeLastStr("12345", "34"))
End Sub

'----------------------------------------
'・LastText
'----------------------------------------
'   ・  大小文字を区別せずに比較する
'----------------------------------------
Public Function IsLastText(ByVal Str, ByVal SubStr)
    IsLastText = IsLastStr(LCase(Str), LCase(SubStr))
End Function

Public Function IncludeLastText(ByVal Str, ByVal SubStr)
    If IsLastText(Str, SubStr) Then
        IncludeLastText = Str
    Else
        IncludeLastText = Str + SubStr
    End If
End Function

Public Function ExcludeLastText(ByVal Str, ByVal SubStr)
    If IsLastText(Str, SubStr) Then
        ExcludeLastText = Mid(Str, 1, Len(Str) - Len(SubStr))
    Else
        ExcludeLastText = Str
    End If
End Function

'----------------------------------------
'・BothStr
'----------------------------------------
Public Function IncludeBothEndsStr(ByVal Str, ByVal SubStr)
    IncludeBothEndsStr = _
        IncludeFirstStr(IncludeLastStr(Str, SubStr), SubStr)
End Function

Public Function ExcludeBothEndsStr(ByVal Str, ByVal SubStr)
    ExcludeBothEndsStr = _
        ExcludeFirstStr(ExcludeLastStr(Str, SubStr), SubStr)
End Function

'----------------------------------------
'・BothText
'----------------------------------------
'   ・  大小文字を区別せずに比較する
'----------------------------------------
Public Function IncludeBothEndsText(ByVal Str, ByVal SubStr)
    IncludeBothEndsText = _
        IncludeFirstText(IncludeLastText(Str, SubStr), SubStr)
End Function

Public Function ExcludeBothEndsText(ByVal Str, ByVal SubStr)
    ExcludeBothEndsText = _
        ExcludeFirstText(ExcludeLastText(Str, SubStr), SubStr)
End Function

'----------------------------------------
'・FirstStrFirst/LastDelim
'----------------------------------------
Public Function FirstStrFirstDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStr(Value, Delimiter)
    If 1 <= Index Then
        Result = Left(Value, Index - 1)
    End If
    FirstStrFirstDelim = Result
End Function

Private Sub testFirstStrFirstDelim
    Call Check("123", FirstStrFirstDelim("123,456", ","))
    Call Check("123", FirstStrFirstDelim("123,456,789", ","))
    Call Check("123", FirstStrFirstDelim("123ttt456", "ttt"))
    Call Check("123", FirstStrFirstDelim("123ttt456", "tt"))
    Call Check("123", FirstStrFirstDelim("123ttt456", "t"))
    Call Check("", FirstStrFirstDelim("123ttt456", ","))
End Sub

'----------------------------------------
'◇First / Last Delim
'----------------------------------------

'----------------------------------------
'・FirstStrLastDelim
'----------------------------------------
Public Function FirstStrLastDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStrRev(Value, Delimiter)
    If 1 <= Index Then
        Result = Left(Value, Index - 1)
    End If
    FirstStrLastDelim = Result
End Function

Private Sub testFirstStrLastDelim
    Call Check("123", FirstStrLastDelim("123,456", ","))
    Call Check("123,456", FirstStrLastDelim("123,456,789", ","))
    Call Check("123", FirstStrLastDelim("123ttt456", "ttt"))
    Call Check("123t", FirstStrLastDelim("123ttt456", "tt"))
    Call Check("123tt", FirstStrLastDelim("123ttt456", "t"))
    Call Check("", FirstStrLastDelim("123ttt456", ","))
End Sub

'----------------------------------------
'・LastStrFirstDelim
'----------------------------------------
Public Function LastStrFirstDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStr(Value, Delimiter)
    If 1 <= Index Then
        Result = Mid(Value, Index + Len(Delimiter))
    End If
    LastStrFirstDelim = Result
End Function

Private Sub testLastStrFirstDelim
    Call Check("456", LastStrFirstDelim("123,456", ","))
    Call Check("456,789", LastStrFirstDelim("123,456,789", ","))
    Call Check("456", LastStrFirstDelim("123ttt456", "ttt"))
    Call Check("t456", LastStrFirstDelim("123ttt456", "tt"))
    Call Check("tt456", LastStrFirstDelim("123ttt456", "t"))
    Call Check("", LastStrFirstDelim("123ttt456", ","))
End Sub

'----------------------------------------
'・LastStrLastDelim
'----------------------------------------
Public Function LastStrLastDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStrRev(Value, Delimiter)
    If 1 <= Index Then
        Result = Mid(Value, Index + Len(Delimiter))
    End If
    LastStrLastDelim = Result
End Function

Private Sub testLastStrLastDelim
    Call Check("456", LastStrLastDelim("123,456", ","))
    Call Check("789", LastStrLastDelim("123,456,789", ","))
    Call Check("456", LastStrLastDelim("123ttt456", "ttt"))
    Call Check("456", LastStrLastDelim("123ttt456", "tt"))
    Call Check("456", LastStrLastDelim("123ttt456", "t"))
    Call Check("", LastStrLastDelim("123ttt456", ","))
End Sub

'----------------------------------------
'◇Trim
'----------------------------------------

Public Function TrimFirstStrs(ByVal Str, ByVal TrimStrs)
    Call Assert(IsArray(TrimStrs), "Error:TrimFirstStrs:TrimStrs is not Array.")
    Dim Result: Result = Str
    Do
        Str = Result
        Dim I
        For I = LBound(TrimStrs) To UBound(TrimStrs)
            Result = ExcludeFirstStr(Result, TrimStrs(I))
        Next
    Loop While Result <> Str
    TrimFirstStrs = Result
End Function

Private Sub testTrimFirstStrs
    Call Check("123 ",              TrimFirstStrs("   123 ", Array(" ")))
    Call Check(vbTab + "  123 ",    TrimFirstStrs("   " + vbTab + "  123 ", Array(" ")))
    Call Check("123 ",              TrimFirstStrs("   " + vbTab + "  123 ", Array(" ", vbTab)))
End Sub

Public Function TrimLastStrs(ByVal Str, ByVal TrimStrs)
    Call Assert(IsArray(TrimStrs), "Error:TrimLastStrs:TrimStrs is not Array.")
    Dim Result: Result = Str
    Do
        Str = Result
        Dim I
        For I = LBound(TrimStrs) To UBound(TrimStrs)
            Result = ExcludeLastStr(Result, TrimStrs(I))
        Next
    Loop While Result <> Str
    TrimLastStrs = Result
End Function

Private Sub testTrimLastStrs
    Call Check(" 123",           TrimLastStrs(" 123   ", Array(" ")))
    Call Check(" 123  " + vbTab, TrimLastStrs(" 123  " + vbTab + "   ", Array(" ")))
    Call Check(" 123",           TrimLastStrs(" 123  " + vbTab + "   ", Array(" ", vbTab)))
End Sub

Public Function TrimBothEndsStrs(ByVal Str, ByVal TrimStrs)
    TrimBothEndsStrs = _
        TrimFirstStrs(TrimLastStrs(Str, TrimStrs), TrimStrs)
End Function

'----------------------------------------
'◇文字列結合
'----------------------------------------

'----------------------------------------
'・文字列結合
'----------------------------------------
'   ・  少なくとも1つのDelimiterが間に入って接続される。
'   ・  Delimiterが結合の両端に付属する場合も1つになる。
'   ・  2連続で結合の両端にある場合は1つが削除される
'       (テストでの動作参照)
'----------------------------------------
Public Function StringCombine(ByVal Delimiter, ByVal Values)
    Call Assert(IsArray(Values), "Error:StringCombine:Values is not Array.")
    Dim Result: Result = ""
    Dim Count: Count = ArrayCount(Values)
    If Count = 0 Then

    ElseIf Count = 1 Then
        Result = Values(0)
    Else
        Dim I
        For I = 0 To Count - 1
        Do
            If Values(I) = "" Then Exit Do
            If Result = "" Then
                Result = Values(I)
            Else
                Result = _
                    ExcludeLastStr(Result, Delimiter) + _
                    Delimiter + _
                    ExcludeFirstStr(Values(I), Delimiter)
            End If

        Loop While False
        Next
    End If
    StringCombine = Result
End Function

Private Sub testStringCombine()
    Call Check("1.2.3.4", StringCombine(".", Array("1", "2", "3", "4")))

    Call Check("1.2", StringCombine(".", Array("1.", "2")))
    Call Check("1.2", StringCombine(".", Array("1.", ".2")))
    Call Check("1..2", StringCombine(".", Array("1..", "2")))
    Call Check("1..2", StringCombine(".", Array("1", "..2")))

    Call Check("1..2", StringCombine(".", Array("1..", ".2")))
    Call Check("1..2", StringCombine(".", Array("1.", "..2")))
    Call Check("1...2", StringCombine(".", Array("1..", "..2")))

    Call Check("1.2.3", StringCombine(".", Array("1.", ".2", "3")))
    Call Check("1.2.3", StringCombine(".", Array("1.", ".2.", "3")))
    Call Check("1.2.3", StringCombine(".", Array("1.", ".2", ".3")))
    Call Check("1.2.3", StringCombine(".", Array("1.", ".2.", ".3")))

    Call Check("1.2.3..4", StringCombine(".", Array("1.", ".2", "3", "..4")))
    Call Check("1..2.3..4", StringCombine(".", Array("1..", ".2", "3.", "..4")))
    Call Check(".1..2.3..4..", StringCombine(".", Array(".1..", ".2", "3.", "..4..")))

    Call Check("", StringCombine(vbCrLf, Array("", "")))
    Call Check("", StringCombine(vbCrLf, Array("", "", "")))
    Call Check("A", StringCombine(vbCrLf, Array("A", "")))
    Call Check("A", StringCombine(vbCrLf, Array("A", "", "")))
    Call Check("A", StringCombine(vbCrLf, Array("", "A")))
    Call Check("A", StringCombine(vbCrLf, Array("", "", "A")))
    Call Check("A", StringCombine(vbCrLf, Array("", "A", "")))

    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("A", "B", "")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("A", "B", "", "")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("", "A", "B")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("", "", "A", "B")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("", "A", "B", "")))

    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("A", "", "B")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("", "A", "", "B")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("A", "", "B", "")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("", "A", "", "B", "")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("A", "", "", "B")))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, Array("", "", "A", "", "", "B", "", "")))

    Call Check("A,B,C", StringCombine(",", Split("A,B,C", ",")))
    Call Check("A,B,C", StringCombine(",", Split("A,B,,C", ",")))
    Call Check("A,B,C", StringCombine(",", Split("A,B,,,C", ",")))
    Call Check("A,B,C", StringCombine(",", Split("A,B,,,,C", ",")))
    Call Check("A,B,C", StringCombine(",", Split(",,A,B,,,,C,,", ",")))

    Call Check("\\test\temp\temp\temp\", StringCombine("\", Array("\\test\", "\temp\", "temp", "\temp\")))
End Sub


'----------------------------------------
'◇文字列比較
'----------------------------------------

'----------------------------------------
'・ワイルドカード検索
'----------------------------------------
'   ・  VBのLike演算子と似た文字列比較
'   ・  厳密には正規表現を利用しているので
'       正規表現文字が含まれていると誤動作する
'----------------------------------------
Public Function LikeCompare(ByVal TargetText, ByVal WildCard)
    Dim Result: Result = False
    Dim RegExp: Set RegExp = CreateObject("VBScript.RegExp")
    WildCard = Replace(WildCard, "*", ".*")
    WildCard = Replace(WildCard, "?", ".")
    WildCard = Replace(WildCard, "\", "\\")
    RegExp.Pattern = IncludeFirstStr(IncludeLastStr(WildCard, "$"), "^")
    Result = RegExp.Test(TargetText)
    LikeCompare = Result
End Function

'----------------------------------------
'・文字列一致を確認する関数
'----------------------------------------
'   ・  部分文字列(キーワード)かワイルドカードで一致確認する
'   ・  ワイルドカード指定かどうかは[*]か[?]が
'       含まれているかどうかで判定する
'----------------------------------------
Public Function MatchText(ByVal TargetText, ByVal SearchStrArray)
    MatchText = (0 <= IndexOfArray(TargetText, SearchStrArray))
End Function

Public Function IndexOfArray(ByVal TargetText, ByVal SearchStrArray)
    Call Assert(IsArray(SearchStrArray), "Error:IndexOfArray:SearchStrArray is not Array.")
    Dim Result: Result = -1
    Dim I
    For I = 0 To ArrayCount(SearchStrArray) - 1
        If IsIncludeStr(SearchStrArray(I), "*") _
        Or IsIncludeStr(SearchStrArray(I), "?")  Then
            'ワイルドカードマッチ
            If (LikeCompare(TargetText, SearchStrArray(I))) Then
                Result = I
                Exit For
            End If
        Else
            'キーワードマッチ
            If (1 <= InStr(TargetText, SearchStrArray(I))) Then
                Result = I
                Exit For
            End If
        End If
    Next
    IndexOfArray = Result
End Function

Private Sub testMatchText
    Call Check(True, MatchText("aaa.ini", Array(".ini")))
    Call Check(False, MatchText("aaa.ini", Array("ab")))
    Call Check(False, MatchText("aaa.ini", Array("ab", "bc")))
    Call Check(True, MatchText("aaa.ini", Array("ab", "bc", "a.i")))
    Call Check(True, MatchText("aaa.ini", Array("*.ini")))
    Call Check(False, MatchText("aaa.ini", Array("*.txt")))
    Call Check(False, MatchText("aaa.ini", Array("*.txt", "123")))
    Call Check(True, MatchText("aaa.ini", Array("*.txt", "123", "*.ini")))
    Call Check(True, MatchText("aaa.ini", Array("*.txt", "123", "a.i")))
End Sub

Public Function MatchTextWildCard(ByVal TargetText, ByVal SearchStrArray)
    MatchTextWildCard = (0 <= IndexOfArrayWildCard(TargetText, SearchStrArray))
End Function

Public Function IndexOfArrayWildCard(ByVal TargetText, ByVal SearchStrArray)
    Call Assert(IsArray(SearchStrArray), "Error:IndexOfArrayWildCard:SearchStrArray is not Array.")
    Dim Result: Result = -1
    Dim I
    For I = 0 To ArrayCount(SearchStrArray) - 1
        'ワイルドカードマッチ
        If (LikeCompare(TargetText, SearchStrArray(I))) Then
            Result = I
            Exit For
        End If
    Next
    IndexOfArrayWildCard = Result
End Function

Private Sub testMatchTextWildCard
    Call Check(False, MatchTextWildCard("aaa.ini", Array(".ini")))
    Call Check(False, MatchTextWildCard("aaa.ini", Array("ab")))
    Call Check(False, MatchTextWildCard("aaa.ini", Array("ab", "bc")))
    Call Check(False, MatchTextWildCard("aaa.ini", Array("ab", "bc", "a.i")))
    Call Check(True, MatchTextWildCard("aaa.ini", Array("*.ini")))
    Call Check(False, MatchTextWildCard("aaa.ini", Array("*.txt")))
    Call Check(False, MatchTextWildCard("aaa.ini", Array("*.txt", "123")))
    Call Check(True, MatchTextWildCard("aaa.ini", Array("*.txt", "123", "*.ini")))
    Call Check(False, MatchTextWildCard("aaa.ini", Array("*.txt", "123", "a.i")))

    Call Check(True, MatchTextWildCard("aaa.ini", Array("aaa.ini")))
End Sub

Public Function MatchTextKeyWord(ByVal TargetText, ByVal SearchStrArray)
    MatchTextKeyWord = (0 <= IndexOfArrayKeyWord(TargetText, SearchStrArray))
End Function

Public Function IndexOfArrayKeyWord(ByVal TargetText, ByVal SearchStrArray)
    Call Assert(IsArray(SearchStrArray), "Error:MatchTextKeyWord:SearchStrArray is not Array.")
    Dim Result: Result = -1
    Dim I
    For I = 0 To ArrayCount(SearchStrArray) - 1
        'キーワードマッチ
        If IsIncludeStr(TargetText, SearchStrArray(I)) Then
            Result = I
            Exit For
        End If
    Next
    IndexOfArrayKeyWord = Result
End Function

Private Sub testMatchTextKeyWord
    Call Check(True, MatchTextKeyWord("aaa.ini", Array(".ini")))
    Call Check(False, MatchTextKeyWord("aaa.ini", Array("ab")))
    Call Check(False, MatchTextKeyWord("aaa.ini", Array("ab", "bc")))
    Call Check(True, MatchTextKeyWord("aaa.ini", Array("ab", "bc", "a.i")))
    Call Check(False, MatchTextKeyWord("aaa.ini", Array("*.ini")))
    Call Check(False, MatchTextKeyWord("aaa.ini", Array("*.txt")))
    Call Check(False, MatchTextKeyWord("aaa.ini", Array("*.txt", "123")))
    Call Check(False, MatchTextKeyWord("aaa.ini", Array("*.txt", "123", "*.ini")))
    Call Check(True, MatchTextKeyWord("aaa.ini", Array("*.txt", "123", "a.i")))
End Sub

'----------------------------------------
'・文字列一致を正規表現で確認する関数
'----------------------------------------
Public Function MatchTextRegExp(ByVal TargetText, ByVal SearchPatternArray)
    MatchTextRegExp = (0 <= IndexOfArrayRegExp(TargetText, SearchPatternArray))
End Function

Public Function IndexOfArrayRegExp(ByVal TargetText, ByVal SearchPatternArray)
    Call Assert(IsArray(SearchPatternArray), "Error:IndexOfArrayRegExp:SearchPatternArray is not Array.")
    Dim RegExp: Set RegExp = CreateObject("VBScript.RegExp")
    Dim Result: Result = -1
    Dim I
    For I = 0 To ArrayCount(SearchPatternArray) - 1
        RegExp.Pattern = SearchPatternArray(I)
        If RegExp.Test(TargetText) Then
            Result = I
            Exit For
        End If
    Next
    IndexOfArrayRegExp = Result
End Function

Private Sub testMatchTextRegExp
    Call Check(True, MatchTextRegExp("2015-03-07", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("20150307", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("_2015-03-07_", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("_20150307_", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?")))
    Call Check(False, MatchTextRegExp("2015/03/07", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("2015/03/07", Array("_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("20150307", Array("_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("_2015/03/07_", Array("_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("_20150307_", Array("_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    Call Check(False, MatchTextRegExp("2015/03/07", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("2015-03-07", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?", "_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("2015/03/07", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?", "_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    Call Check(True, MatchTextRegExp("20150307", Array("_?\d{4}-?\d{1,2}-?\d{1,2}_?", "_?\d{4}/?\d{1,2}/?\d{1,2}_?")))
    'Call MsgBox("testMatchRegExp")
End Sub


'----------------------------------------
'◇文字列置換
'----------------------------------------

'----------------------------------------
'・正規表現に一致するものを削除する関数
'----------------------------------------
Public Function DeleteRegText(ByVal Value, ByVal Pattern)
    Dim Result: Result = Value
    Dim RegExp: Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Pattern = Pattern
    Do While RegExp.Test(Result)
        Result = RegExp.Replace(Result, "")
    Loop
    'パターンに一致し続ける限りは置き換えを行う
    DeleteRegText = Result
End Function

'----------------------------------------
'・日付文字列を削除する関数
'----------------------------------------
'   ・  [_YYYY-MM-DD_]や[YYYYMMDD]という文字列を
'       元の文字列から削除する
'----------------------------------------
Public Function DeleteDateText(ByVal Value)
'    DeleteDateText = DeleteRegText(Value, _
'        "_?\d{4}-?\d{1,2}-?\d{1,2}_?")
    DeleteDateText = DeleteRegText(Value, _
        "(^\d{4}-\d{1,2}-\d{1,2}_?|" + _
        "^\d{4}\d{2}\d{2}_?|" + _
        "_?\d{4}-\d{1,2}-\d{1,2}$|" + _
        "_?\d{4}\d{2}\d{2}$)")
End Function

Private Sub testDeleteDateText
    Call Check("ABCDEFG", DeleteDateText("2015-03-06_ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateText("ABCDEFG_2015-03-06"))
    Call Check("ABCDEFG", DeleteDateText("2015-03-06_ABCDEFG_2015-03-06"))

    Call Check("ABCDEFG", DeleteDateText("2015-03-06ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateText("ABCDEFG2015-03-06"))
    Call Check("ABCDEFG", DeleteDateText("2015-03-06ABCDEFG2015-03-06"))

    Call Check("ABCDEFG", DeleteDateText("20150306_ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateText("ABCDEFG_20150306"))
    Call Check("ABCDEFG", DeleteDateText("20150306_ABCDEFG_20150306"))

    Call Check("ABCDEFG", DeleteDateText("20150306ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateText("ABCDEFG20150306"))
    Call Check("ABCDEFG", DeleteDateText("20150306ABCDEFG20150306"))

    'Call WScript.Echo("Test Finish")
End Sub

'----------------------------------------
'・日付時刻文字列を削除する関数
'----------------------------------------
'   ・  [_YYYY-MM-DD_HH-MM-DD_]や[YYYYMMDDHHMMDD]という文字列を
'       元の文字列から削除する
'----------------------------------------
Public Function DeleteDateTimeText(ByVal Value)
'    DeleteDateTimeText = DeleteRegText(Value, _
'        "_?\d{4}-?\d{1,2}-?\d{1,2}_?-?\d{1,2}-?\d{1,2}-?\d{1,2}_?")
    DeleteDateTimeText = DeleteRegText(Value, _
        "(^\d{4}-\d{1,2}-\d{1,2}_\d{1,2}-\d{1,2}-\d{1,2}_?|" + _
        "^\d{4}\d{2}\d{2}(-|_)?\d{2}\d{2}\d{2}_?|" + _
        "_?\d{4}-\d{1,2}-\d{1,2}_\d{1,2}-\d{1,2}-\d{1,2}$|" + _
        "_?\d{4}\d{2}\d{2}(-|_)?\d{2}\d{2}\d{2})")
End Function

Private Sub testDeleteDateTimeText
    Call Check("ABCDEFG", DeleteDateTimeText("2015-03-06_22-25-57_ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateTimeText("ABCDEFG_2015-03-06_22-25-57"))
    Call Check("ABCDEFG", DeleteDateTimeText("2015-03-06_22-25-57_ABCDEFG_2015-03-06_22-25-57"))

    Call Check("ABCDEFG", DeleteDateTimeText("2015-03-06_22-25-57ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateTimeText("ABCDEFG2015-03-06_22-25-57"))
    Call Check("ABCDEFG", DeleteDateTimeText("2015-03-06_22-25-57ABCDEFG2015-03-06_22-25-57"))

    Call Check("ABCDEFG", DeleteDateTimeText("20150306_222557_ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateTimeText("ABCDEFG_20150306_222557"))
    Call Check("ABCDEFG", DeleteDateTimeText("20150306_222557_ABCDEFG_20150306_222557"))

    Call Check("ABCDEFG", DeleteDateTimeText("20150306-222557_ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateTimeText("ABCDEFG_20150306-222557"))
    Call Check("ABCDEFG", DeleteDateTimeText("20150306-222557_ABCDEFG_20150306-222557"))

    Call Check("ABCDEFG", DeleteDateTimeText("20150306222557ABCDEFG"))
    Call Check("ABCDEFG", DeleteDateTimeText("ABCDEFG20150306222557"))
    Call Check("ABCDEFG", DeleteDateTimeText("20150306222557ABCDEFG20150306222557"))


'    Call Check("ABCDEFG", DeleteDateTimeText("2015-03-06-22-25-57_ABC2015-03-06_22-25-57DEFG_2015-03-06-22-25-57"))
'    Call Check("ABCDEFG", DeleteDateTimeText("ABC_2015-03-06_22-25-57_DEFG"))
'    Call Check("ABCDEFG", DeleteDateTimeText("ABC_201503-06-222557_DEFG"))
'    Call Check("ABCDEFG", DeleteDateTimeText("ABC_20150306222557_DEFG"))
'    Call Check("ABCDEFG", DeleteDateTimeText("ABC2015030622-25-57DE20150306222557FG"))
    'Call WScript.Echo("Test Finish")
End Sub


'----------------------------------------
'◇標準文字列関数動作確認テスト
'----------------------------------------

'----------------------------------------
'・Split
'----------------------------------------
Private Sub testSplit
    Call Check(0, ArrayCount(Split("", "-")))
    Call Check(1, ArrayCount(Split("aaa", "-")))
    Call Check(2, ArrayCount(Split("aaa-", "-")))
    Call Check(2, ArrayCount(Split("-aaa", "-")))
    Call Check(2, ArrayCount(Split("a-aa", "-")))
    Call Check(3, ArrayCount(Split("a-a-a", "-")))
    Call Check(4, ArrayCount(Split("-a-a-a", "-")))
    Call Check(4, ArrayCount(Split("a-a-a-", "-")))
    Call Check(5, ArrayCount(Split("-a-a-a-", "-")))
End Sub

'----------------------------------------
'◆配列処理
'----------------------------------------
'VBScriptは
'   Dim A(3 To 5)
'   ReDim B(4 To 6)
'というような宣言ができないので
'LBoundは常に0だと考えて処理を実装する
'例外はあるかもしれないが不明
'----------------------------------------

'----------------------------------------
'・配列の要素数を求める関数
'----------------------------------------
Public Function ArrayCount(ByVal ArrayValue)
    Call Assert(IsArray(ArrayValue), "Error:ArrayCount:ArrayValue is not Array.")
    ArrayCount = UBound(ArrayValue) - LBound(ArrayValue) + 1
End Function

Private Sub testArrayCount
    Dim A
    A = Array("A", "B", "C")
    Call Check(3, ArrayCount(A))

    ReDim Preserve A(4) '0-4:Count=5
    Call Check(5, ArrayCount(A))
End Sub

'----------------------------------------
'・配列の要素を追加する
'----------------------------------------
'   ・  オブジェクト値にも対応
'----------------------------------------
Sub ArrayAdd(ByRef ArrayValue, ByVal Value)
    Call Assert(IsArray(ArrayValue), "Error:ArrayAdd:ArrayValue is not Array.")

    ReDim Preserve ArrayValue(ArrayCount(ArrayValue))
    Call SetValue(ArrayValue(UBound(ArrayValue)), Value)
End Sub

Private Sub testArrayAdd
    Dim A
    A = Array("A", "B", "C")

    Call ArrayAdd(A, "D")
    Call Check(4, ArrayCount(A))
    Call Check("D", A(3))

    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = CreateObject("ADODB.Stream")
    Call ArrayAdd(B, fso)
    Call Check("test.txt", B(3).GetFileName("C:\temp\test.txt"))
End Sub

'----------------------------------------
'・配列の要素を挿入する
'----------------------------------------
'   ・  オブジェクト値にも対応
'----------------------------------------
Sub ArrayInsert(ByRef ArrayValue, ByVal Index, ByVal Value)
    Call Assert(IsArray(ArrayValue), "Error:ArrayInsert:ArrayValue is not Array.")
    Call Assert(InRange(LBound(ArrayValue), Index, UBound(ArrayValue)), _
        "Error:ArrayInsert:Index Range Over.")

    ReDim Preserve ArrayValue(UBound(ArrayValue) + 1)
    Dim I
    For I = UBound(ArrayValue) To Index + 1 Step -1
        Call SetValue(ArrayValue(I), ArrayValue(I - 1))
    Next
    Call SetValue(ArrayValue(Index), Value)
End Sub

Private Sub testArrayInsert
    Dim A
    A = Array("A", "B", "C")

    Call Check("B", A(1))
    Call Check(3, ArrayCount(A))
    Call ArrayInsert(A, 1, "1")
    Call Check(4, ArrayCount(A))
    Call Check("1", A(1))

    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = CreateObject("ADODB.Stream")
    Call Check(CurrentDirectory, B(1).CurrentDirectory)
    Call ArrayInsert(B, 1, fso)
    Call Check("test.txt", B(1).GetFileName("C:\temp\test.txt"))
End Sub

'----------------------------------------
'・配列の要素を削除する
'----------------------------------------
'   ・  オブジェクト値にも対応
'----------------------------------------
Sub ArrayDelete(ByRef ArrayValue, ByVal Index)
    Call Assert(IsArray(ArrayValue), "Error:ArrayDelete:ArrayValue is not Array.")
    Call Assert(InRange(LBound(ArrayValue), Index, UBound(ArrayValue)), _
        "Error:ArrayDelete:Index Range Over.")

    Dim I
    For I = Index + 1 To UBound(ArrayValue)
        Call SetValue(ArrayValue(I - 1), ArrayValue(I))
    Next
    ReDim Preserve ArrayValue(UBound(ArrayValue) - 1)
End Sub

Private Sub testArrayDelete
    Dim A
    A = Array("A", "B", "C")

    Call Check("B", A(1))
    Call ArrayDelete(A, 1)
    Call Check(2, ArrayCount(A))
    Call Check("C", A(1))

    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = fso
    Call Check(CurrentDirectory, B(1).CurrentDirectory)
    Call ArrayDelete(B, 1)
    Call Check("test.txt", B(1).GetFileName("C:\temp\test.txt"))
End Sub

'----------------------------------------
'・配列関数のテスト
'----------------------------------------
Sub testArrayFunctions
    Dim A
    A = Array("A", "B", "C")
    Call Check(3, ArrayCount(A))
    Call Check("A B C", ArrayToString(A, " "))
    Call ArrayAdd(A, "D")
    Call Check(4, ArrayCount(A))
    Call Check("A B C D", ArrayToString(A, " "))

    Call ArrayDelete(A, 0)
    Call Check("B C D", ArrayToString(A, " "))
    Call ArrayDelete(A, 2)
    Call Check("B C", ArrayToString(A, " "))

    A = Array("A", "B", "C")
    Call ArrayDelete(A, 1)
    Call Check("A C", ArrayToString(A, " "))

    Call ArrayInsert(A, 1, "B")
    Call Check("A B C", ArrayToString(A, " "))
    Call ArrayInsert(A, 0, "1")
    Call Check("1 A B C", ArrayToString(A, " "))
    Call ArrayInsert(A, 3, "2")
    Call Check("1 A B 2 C", ArrayToString(A, " "))
End Sub

'----------------------------------------
'・配列を文字列にして単純に結合する関数
'----------------------------------------
Function ArrayToString(ByRef ArrayValue, ByVal Delimiter)
    Call Assert(IsArray(ArrayValue), "配列ではありません")
    Dim Result: Result = ""
    Dim I
    For I = LBound(ArrayValue) To UBound(ArrayValue)
        Result = Result + ArrayValue(I) + Delimiter
    Next
    Result = ExcludeLastStr(Result, Delimiter)
    ArrayToString = Result
End Function

'----------------------------------------
'・配列同士を結合する関数
'----------------------------------------
Sub ArrayAddArray(ByRef ArrayValue, ByVal AddArrayValue)
    Call Assert(IsArray(ArrayValue), "Error:ArrayAddArray:ArrayValue is not array.")
    Call Assert(IsArray(AddArrayValue), "Error:ArrayAddArray:AddArrayValue is not array.")

    Dim I
    For I = LBound(AddArrayValue) To UBound(AddArrayValue)
        Call ArrayAdd(ArrayValue, AddArrayValue(I))
    Next
End Sub


'----------------------------------------
'◆日付時刻処理
'----------------------------------------

'----------------------------------------
'・日付書式
'----------------------------------------
Public Function FormatYYYYMMDD(ByVal DateValue)
    FormatYYYYMMDD = FormatYYYY_MM_DD(DateValue, "")
End Function

Sub testFormatYYYYMMDD
    Dim Value: Value = CDate("2015/02/03")
    Call Check("20150203", FormatYYYYMMDD(Value))
End Sub

Public Function FormatYYYY_MM_DD(ByVal DateValue, ByVal Delimiter)
    Dim Result: Result = ""
    Do
        If IsDate(DateValue) = False Then Exit Do

        Result = _
            Year(DateValue) & _
            Delimiter & _
            Right("0" & Month(DateValue), 2) & _
            Delimiter & _
            Right("0" & Day(DateValue), 2)
    Loop While False
    FormatYYYY_MM_DD = Result
End Function

'----------------------------------------
'・時刻書式
'----------------------------------------
Public Function FormatHHMMSS(ByVal TimeValue)
    FormatHHMMSS = FormatHH_MM_SS(TimeValue, "")
End Function

Sub testFormatHHMMSS
    Dim Value: Value = CDate("2015/02/03 05:05")
    Call Check("05:05:00", FormatHH_MM_SS(Value, ":"))
End Sub

Public Function FormatHH_MM_SS(ByVal TimeValue, ByVal Delimiter)
    Dim Result: Result = ""
    Do
        If IsDate(TimeValue) = False Then Exit Do

        Result = _
            Right("0" & Hour(TimeValue) , 2) & _
            Delimiter & _
            Right("0" & Minute(TimeValue) , 2) & _
            Delimiter & _
            Right("0" & Second(TimeValue) , 2)
    Loop While False
    FormatHH_MM_SS = Result
End Function

'----------------------------------------
'・日付時刻書式
'----------------------------------------
Public Function FormatYYYYMMDDHHMMSS(ByVal DateTimeValue)
    FormatYYYYMMDDHHMMSS = _
        FormatYYYYMMDD(DateTimeValue) + _
        FormatHHMMSS(DateTimeValue)
End Function

'----------------------------------------
'◇書式指定
'----------------------------------------
'Const vbSunday = 1
'Const vbMonday = 2
'Const vbTuesday = 3
'Const vbWednesday = 4
'Const vbThursday = 5
'Const vbFriday = 6
'Const vbSaturday = 7

Public Function WeekDayJapaneseShort(ByVal DateValue)
    Dim Result: Result = ""
    Do
        If IsDate(DateValue) = False Then Exit Do

        Select Case Weekday(DateValue)
        Case vbSunday:      Result = "日"
        Case vbMonday:      Result = "月"
        Case vbTuesday:     Result = "火"
        Case vbWednesday:   Result = "水"
        Case vbThursday:    Result = "木"
        Case vbFriday:      Result = "金"
        Case vbSaturday:    Result = "土"
        End Select
    Loop While False
    WeekDayJapaneseShort = Result
End Function

Public Function WeekDayJapaneseLong(ByVal DateValue)
    Dim Result: Result = ""
    Do
        If IsDate(DateValue) = False Then Exit Do
        Result = WeekDayJapaneseShort(DateValue) + "曜日"
    Loop While False
    WeekDayJapaneseLong = Result
End Function

Public Function WeekDayEnglishShort(ByVal DateValue)
    Dim Result: Result = ""
    Do
        If IsDate(DateValue) = False Then Exit Do

        Select Case Weekday(DateValue)
        Case vbSunday:      Result = "Sun"
        Case vbMonday:      Result = "Mon"
        Case vbTuesday:     Result = "Tue"
        Case vbWednesday:   Result = "Wed"
        Case vbThursday:    Result = "Thu"
        Case vbFriday:      Result = "Fri"
        Case vbSaturday:    Result = "Sat"
        End Select
    Loop While False
    WeekDayEnglishShort = Result
End Function

Public Function WeekDayEnglishLong(ByVal DateValue)
    Dim Result: Result = ""
    Do
        If IsDate(DateValue) = False Then Exit Do

        Select Case Weekday(DateValue)
        Case vbSunday:      Result = "Sunday"
        Case vbMonday:      Result = "Monday"
        Case vbTuesday:     Result = "Tuesday"
        Case vbWednesday:   Result = "Wednesday"
        Case vbThursday:    Result = "Thursday"
        Case vbFriday:      Result = "Friday"
        Case vbSaturday:    Result = "Saturday"
        End Select
    Loop While False
    WeekDayEnglishLong = Result
End Function
'----------------------------------------
'・書式日付時刻変換関数
'----------------------------------------
'   ・  対応書式文字列
'   yyyy/yy/y
'   mm/m
'   dd/d
'   dddd/ddd    Sunday..Saturday / Sun..Sat
'   aaaa/aaa    日曜日..土曜日 / 日..土
'   hh/h
'   nn/n
'   ss/s
Public Function FormatDateTimeToString(ByVal DateValue, ByVal FormatStr)
    Dim Result: Result = ""
    Do
        If IsDate(DateValue) = False Then Exit Do

        Result = LCase(FormatStr)

        Result = Replace(Result, "yyyy", Year(DateValue))
        Result = Replace(Result, "yy", Right(Year(DateValue), 2))
        Result = Replace(Result, "y", Right(Year(DateValue), 1))

        Result = Replace(Result, "mm", Right("0" & Month(DateValue), 2))
        Result = Replace(Result, "m", Month(DateValue))

        Result = Replace(Result, "dddd", "DDDD")
        Result = Replace(Result, "ddd", "DDD")

        Result = Replace(Result, "dd", Right("0" & Day(DateValue), 2))
        Result = Replace(Result, "d", Day(DateValue))

        Result = Replace(Result, "aaaa", WeekDayJapaneseLong(DateValue))
        Result = Replace(Result, "aaa", WeekDayJapaneseShort(DateValue))

        Result = Replace(Result, "hh", Right("0" & Hour(DateValue), 2))
        Result = Replace(Result, "h", Hour(DateValue))

        Result = Replace(Result, "nn", Right("0" & Minute(DateValue), 2))
        Result = Replace(Result, "n", Minute(DateValue))

        Result = Replace(Result, "ss", Right("0" & Second(DateValue), 2))
        Result = Replace(Result, "s", Second(DateValue))

        Result = Replace(Result, "DDDD", WeekDayEnglishLong(DateValue))
        Result = Replace(Result, "DDD", WeekDayEnglishShort(DateValue))

    Loop While False
    FormatDateTimeToString = Result
End Function

Sub testFormatDateTimeToString
    Dim Value
    Value = CDate("2015/01/02 3:4:5")
    Call Check("2015", FormatDateTimeToString(Value, "YYYY"))
    Call Check("2015", FormatDateTimeToString(Value, "yyyy"))
    Call Check("15", FormatDateTimeToString(Value, "yY"))
    Call Check("5", FormatDateTimeToString(Value, "y"))
    Call Check("5", FormatDateTimeToString(Value, "y"))
    Call Check("01", FormatDateTimeToString(Value, "mm"))
    Call Check("1", FormatDateTimeToString(Value, "m"))
    Call Check("02", FormatDateTimeToString(Value, "dd"))
    Call Check("2", FormatDateTimeToString(Value, "d"))
    Call Check("03", FormatDateTimeToString(Value, "hh"))
    Call Check("3", FormatDateTimeToString(Value, "h"))
    Call Check("04", FormatDateTimeToString(Value, "nn"))
    Call Check("4", FormatDateTimeToString(Value, "n"))
    Call Check("05", FormatDateTimeToString(Value, "ss"))
    Call Check("5", FormatDateTimeToString(Value, "s"))

    Call Check("金", FormatDateTimeToString(Value, "aaa"))
    Call Check("金曜日", FormatDateTimeToString(Value, "aaaa"))
    Call Check("Fri", FormatDateTimeToString(Value, "ddd"))
    Call Check("Friday", FormatDateTimeToString(Value, "dddd"))

    Value = CDate("2015/11/12 13:14:15")
    Call Check("11", FormatDateTimeToString(Value, "mm"))
    Call Check("11", FormatDateTimeToString(Value, "m"))
    Call Check("12", FormatDateTimeToString(Value, "dd"))
    Call Check("12", FormatDateTimeToString(Value, "d"))
    Call Check("13", FormatDateTimeToString(Value, "hh"))
    Call Check("13", FormatDateTimeToString(Value, "h"))
    Call Check("14", FormatDateTimeToString(Value, "nn"))
    Call Check("14", FormatDateTimeToString(Value, "n"))
    Call Check("15", FormatDateTimeToString(Value, "ss"))
    Call Check("15", FormatDateTimeToString(Value, "s"))
End Sub

'----------------------------------------
'◆ファイルフォルダパス処理
'----------------------------------------

'----------------------------------------
'・絶対パスを取得する関数
'カレントディレクトリパスと相対パスを指定する
'----------------------------------------
Public Function AbsoluteFilePath(ByVal BasePath, ByVal RelativePath)
    AbsoluteFilePath = AbsolutePath(BasePath, RelativePath)
End Function

Public Function AbsolutePath(ByVal BasePath, ByVal RelativePath)
    Dim Result
    Do
        If fso.FolderExists(BasePath) = False Then
            Result = RelativePath
            Exit Do
        End If

        '相対パス指定部分が絶対パスの場合、
        '処理は行わずにそのまま絶対パスを返す
        If IsIncludeDrivePath(RelativePath) Then
            Result = RelativePath
            Exit Do
        End If

        '相対パス指定部分がネットワークパスの場合、
        '処理は行わずにそのまま絶対パスを返す
        If IsNetworkPath(RelativePath) Then
            Result = RelativePath
            Exit Do
        End If

        'カレントディレクトリを確保
        Dim CurDirBuffer
        CurDirBuffer = Shell.CurrentDirectory

        Shell.CurrentDirectory = BasePath

        '相対パスを取得
        Result = fso.GetAbsolutePathName(RelativePath)

        'カレントディレクトリを元に戻す
        Shell.CurrentDirectory =  CurDirBuffer

    Loop While False
    AbsolutePath = Result
End Function

Sub testAbsolutePath()

     '通常の相対パス指定
    Call Check(LCase("C:\Windows\System32"), LCase(AbsolutePath("C:\Windows", ".\System32")))
    Call Check(LCase("C:\Windows\System32"), LCase(AbsolutePath("C:\Program Files", "..\Windows\System32")))

    'ピリオドではない相対パス指定
    Call Check(LCase("C:\Windows\System32"), LCase(AbsolutePath("C:\Windows", "System32")))

    'ドライブパス指定
    Call Check(LCase("C:\Program Files"), LCase(AbsolutePath("C:\Windows", "C:\Program Files")))
    Call Check(LCase("C:\Windows\System32"), LCase(AbsolutePath("C:\Program Files", "C:\Windows\System32")))

    'ネットワークパス指定
    Call Check(LCase("\\127.0.0.1\C$"), LCase(AbsolutePath("C:\Windows", "\\127.0.0.1\C$")))

    Call Check(AbsolutePath(ScriptFolderPath, ".\Test\TestFileFolderPathList"), _
        ScriptFolderPath + "\Test\TestFileFolderPathList")
    Call Check(AbsolutePath(ScriptFolderPath, "..\Test\TestFileFolderPathList"), _
        fso.GetParentFolderName(ScriptFolderPath) + "\Test\TestFileFolderPathList")

    MsgBox "Test OK"
End Sub

'----------------------------------------
'・ドライブパスが含まれているかどうか確認する関数
'[:]が2文字目以降にあるかどうかで判定
'----------------------------------------
Public Function IsIncludeDrivePath(ByVal Path)
    Dim Result
    Result = (2 <= InStr(Path, ":"))
    IsIncludeDrivePath = Result
End Function
'
'----------------------------------------
'・ネットワークドライブかどうか確認する関数
'----------------------------------------
Public Function IsNetworkPath(ByVal Path)
    Dim Result: Result = False
    If IsFirstStr(Path, "\\") Then
        If 3 <= Len(Path) Then
            Result = True
        End If
    End If
    IsNetworkPath = Result
End Function
'
'----------------------------------------
'・ドライブパス"C:"を取り出す関数
'----------------------------------------
Public Function GetDrivePath(ByVal Path)
    Dim Result: Result = ""
    If IsIncludeDrivePath(Path) Then
        Result = FirstStrFirstDelim(Path, ":")
        Result = IncludeLastStr(Result, ":")
    End If
    GetDrivePath = Result
End Function

'----------------------------------------
'・終端にパス区切りを追加する関数
'----------------------------------------
Public Function IncludeLastPathDelim(ByVal Path)
    Dim Result: Result = ""
    If Path <> "" Then
        Result = IncludeLastStr(Path, "\")
    End If
    IncludeLastPathDelim = Result
End Function

'----------------------------------------
'・終端からパス区切りを削除する関数
'----------------------------------------
Public Function ExcludeLastPathDelim(ByVal Path)
    Dim Result: Result = ""
    If Path <> "" Then
        Result = ExcludeLastStr(Path, "\")
    End If
    ExcludeLastPathDelim = Result
End Function

'----------------------------------------
'・スペースの含まれた値をダブルクウォートで囲う
'----------------------------------------
Function InSpacePlusDoubleQuote(ByVal Value)
    Dim Result: Result = ""
    If 0 < InStr(Value, " ") Then
        Result = IncludeBothEndsStr(Value, """")
    Else
        Result = Value
    End If
    InSpacePlusDoubleQuote = Result
End Function

'----------------------------------------
'・ファイルパス配列をダブルクウォートで囲み指定文字で連結する
'----------------------------------------
Function IncludeBothEndsDoubleQuoteCombineArray(ByVal ArrayValue, ByVal Delimiter)
    Call Assert(IsArray(ArrayValue), _
        "Error:Function IncludeBothEndsDoubleQuoteCombineArray:ArrayValue is not array")

    Dim Result: Result = ""
    Dim I
    For I = 0 To ArrayCount(ArrayValue) - 1
        ArrayValue(I) = IncludeBothEndsStr(ArrayValue(I), """")
    Next
    Result = StringCombine(Delimiter, ArrayValue)

    IncludeBothEndsDoubleQuoteCombineArray = Result
End Function


'----------------------------------------
'・ピリオドを含む拡張子を取得する関数
'----------------------------------------
'ピリオドのないファイル名の場合は空文字を返す
'fso.GetExtensionName ではピリオドで終わるファイル名を
'判断できないために作成した。
'----------------------------------------
Public Function PeriodExtName(ByVal Path)
    Dim Result: Result = ""
    Result = fso.GetExtensionName(Path)
    If IsLastStr(Path, "." + Result) Then
        Result = "." + Result
    End If
    PeriodExtName = Result
End Function

Sub testPeriodExtName
    Call Check(".txt", PeriodExtName("C:\temp\test.txt"))
    Call Check(".zip", PeriodExtName("C:\temp\test.tmp.zip"))
    Call Check("", PeriodExtName("C:\temp\test"))
    Call Check(".", PeriodExtName("C:\temp\test."))
End Sub

'----------------------------------------
'・ファイル名から拡張子を取り除く関数
'----------------------------------------
Public Function ExcludePathExt(ByVal Path)
    Dim Result: Result = ""
    If Path <> "" Then
        Result = ExcludeLastStr(Path, PeriodExtName(Path))
    End If
    ExcludePathExt = Result
End Function

Sub testExcludePathExt
    Call Check("C:\temp\test",      ExcludePathExt("C:\temp\test.txt"))
    Call Check("C:\temp\test.tmp",  ExcludePathExt("C:\temp\test.tmp.zip"))
    Call Check("C:\temp\test",      ExcludePathExt("C:\temp\test"))
    Call Check("C:\temp\test",      ExcludePathExt("C:\temp\test."))
End Sub

'----------------------------------------
'・ファイルパスの拡張子を変更する関数
'----------------------------------------
Public Function ChangeFileExt(ByVal Path, ByVal NewExt)
    Dim Result: Result = ""
    If Path <> "" Then
        NewExt = IIF(NewExt = "", "", IncludeFirstStr(NewExt, "."))
        Result = ExcludePathExt(Path) + NewExt
    End If
    ChangeFileExt = Result
End Function

Sub testChangeFileExt
    Call Check("C:\temp\test.zip",      ChangeFileExt("C:\temp\test.txt", "zip"))
    Call Check("C:\temp\test.7z.tmp",   ChangeFileExt("C:\temp\test.txt", ".7z.tmp"))
    Call Check("C:\temp\test.tmp.7z",   ChangeFileExt("C:\temp\test.tmp.zip", "7z"))
    Call Check("C:\temp\test.",         ChangeFileExt("C:\temp\test", "."))
    Call Check("C:\temp\test",          ChangeFileExt("C:\temp\test.", ""))
End Sub

'----------------------------------------
'・ファイルパスを結合する関数
'----------------------------------------
Public Function PathCombine(ByVal Values)
    Call Assert(IsArray(Values), "Error:PathCombine")
    PathCombine = StringCombine("\", Values)
End Function

Sub testPathCombine()
    Call Check("C:\Temp\Temp\temp.txt", PathCombine(Array("C:", "Temp", "Temp", "temp.txt")))
    Call Check("C:\Temp\Temp\temp.txt", PathCombine(Array("C:\Temp", "Temp\temp.txt")))
    Call Check("C:\Temp\Temp\temp.txt", PathCombine(Array("C:\Temp\Temp\", "\temp.txt")))
    Call Check("\Temp\Temp\", PathCombine(Array("\Temp\", "\Temp\")))

    Call Check("C:\work\bbb\a.txt", PathCombine(Array("C:\work",    "bbb\a.txt")))
    Call Check("C:\work\bbb\a.txt", PathCombine(Array("C:\work\",   "bbb\a.txt")))
    Call Check("C:\work\bbb\a.txt", PathCombine(Array("C:\work",    "\bbb\a.txt")))
    Call Check("C:\work\bbb\a.txt", PathCombine(Array("C:\work\",   "\bbb\a.txt")))

    Call Check("C:\work\bbb\a.txt", PathCombine(Array("C:\work",    "bbb",  "a.txt")))
    Call Check("C:\work\bbb\a.txt", PathCombine(Array("C:\work\",   "\bbb\","\a.txt")))

    Call Check("\C:\work\bbb\a.txt\", PathCombine(Array("\C:\work\",  "\bbb\","\a.txt\")))

End Sub

'----------------------------------------
'◆ファイルフォルダパス取得
'----------------------------------------

'----------------------------------------
'・カレントディレクトリの取得
'----------------------------------------
Public Function CurrentDirectory
    CurrentDirectory = Shell.CurrentDirectory
End Function

'----------------------------------------
'・スクリプトフォルダの取得
'----------------------------------------
Public Function ScriptFolderPath
    ScriptFolderPath = _
        fso.GetParentFolderName(WScript.ScriptFullName)
End Function

'----------------------------------------
'・一時ファイルの取得
'----------------------------------------
'   ・  このようなパスが取得できる
'       C:\Users\<UserName>\AppData\Local\Temp\rad92218.tmp
'----------------------------------------
Public Function TemporaryPath
    TemporaryPath = TemporaryFilePath
End Function

Public Function TemporaryFilePath
    Dim Result
    Const TemporaryFolder = 2

    ' リダイレクト先のファイル名を生成。
    Do
        Result = fso.BuildPath( _
            fso.GetSpecialFolder(TemporaryFolder).Path, _
            fso.GetTempName)
    Loop While fso.FileExists(Result)
    TemporaryFilePath = Result
End Function

'----------------------------------------
'・一時フォルダの取得
'----------------------------------------
'   ・  このようなパスが取得できる
'       C:\Users\<UserName>\AppData\Local\Temp\rad92218
'----------------------------------------
Public Function TemporaryFolderPath
    Dim Result
    Do
        Result = ChangeFileExt(TemporaryFilePath, "")
    Loop While fso.FolderExists(Result)
    TemporaryFolderPath = Result
End Function

'----------------------------------------
'・一時ファイル名
'----------------------------------------
'   ・  このような名前が取得できる
'       [rad92218.tmp]
'----------------------------------------
Function TemporaryFileName
    TemporaryFileName = fso.GetTempName
End Function


'----------------------------------------
'・一時フォルダ名
'----------------------------------------
'   ・  このような名前が取得できる
'       [rad92218]
'----------------------------------------
Function TemporaryFolderName
    TemporaryFolderName = ChangeFileExt(fso.GetTempName, "")
End Function

'----------------------------------------
'◆ファイルフォルダ列挙
'----------------------------------------

Sub testFileFolderPathList()
    Dim Path: Path = AbsolutePath(ScriptFolderPath, ".\Test\TestFileFolderPathList")
    Dim PathList

    PathList = Replace(UCase(FolderPathListTopFolder(Path)), UCase(Path), "")
    'Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, Array( _
            "\AAA", _
            "\BBB" _
        )))

    PathList = Replace(UCase(FolderPathListSubFolder(Path)), UCase(Path), "")
    'Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, Array( _
            "\AAA", _
            "\AAA\AAA-1", _
            "\AAA\AAA-2", _
            "\AAA\AAA-2\AAA-2-1", _
            "\AAA\AAA-2\AAA-2-2", _
            "\AAA\AAA-2\AAA-2-2\AAA-2-2-1", _
            "\BBB", _
            "\BBB\BBB-1", _
            "\BBB\BBB-1\BBB-1-1", _
            "\BBB\BBB-1\BBB-1-1\BBB-1-1-1", _
            "\BBB\BBB-1\BBB-1-2", _
            "\BBB\BBB-2", "" _
        )))

    PathList = Replace(UCase(FilePathListTopFolder(Path)), UCase(Path), "")
    'Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, Array( _
            "\AAA.TXT", _
            "\BBB.TXT" _
        )))

    PathList = Replace(UCase(FilePathListSubFolder(Path)), UCase(Path), "")
    'Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, Array( _
            "\AAA\AAA-1.TXT", _
            "\AAA\AAA-2.TXT", _
            "\AAA\AAA-1\AAA-1-1.TXT", _
            "\AAA\AAA-2\AAA-2-1.TXT", _
            "\AAA\AAA-2\AAA-2-2.TXT", _
            "\AAA\AAA-2\AAA-2-1\AAA-2-1-1.TXT", _
            "\AAA\AAA-2\AAA-2-2\AAA-2-2-1.TXT", _
            "\AAA\AAA-2\AAA-2-2\AAA-2-2-1\AAA-2-2-1-1.TXT", _
            "\BBB\BBB-1.TXT", _
            "\BBB\BBB-2.TXT", _
            "\BBB\BBB-1\BBB-1-1.TXT", _
            "\BBB\BBB-1\BBB-1-2.TXT", _
            "\BBB\BBB-1\BBB-1-1\BBB-1-1-1.TXT", _
            "\BBB\BBB-1\BBB-1-1\BBB-1-1-1\BBB-1-1-1-1.TXT", _
            "\BBB\BBB-1\BBB-1-2\BBB-1-2-1.TXT", _
            "\BBB\BBB-2\BBB-2-1.TXT", _
            "\AAA.TXT", _
            "\BBB.TXT" _
        )))
End Sub

'----------------------------------------
'・フォルダリストを取得
'----------------------------------------
'   ・  存在しなければ空文字を返す。
'   ・  パスは改行コードで区切られている
'----------------------------------------
Public Function FolderPathListTopFolder(ByVal FolderPath)
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FolderPathListTopFolder:Folder no Exists")
    Dim Result: Result = ""
    Dim SubFolder
    For Each SubFolder In fso.GetFolder(FolderPath).SubFolders
        Result = StringCombine(vbCrLf, Array(Result, SubFolder.Path))
    Next
    FolderPathListTopFolder = Result
End Function

Public Function FolderPathListSubFolder(ByVal FolderPath)
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FolderPathListSubFolder:Folder no Exists")
    Dim Result: Result = ""
    Dim SubFolder
    For Each SubFolder In fso.GetFolder(FolderPath).SubFolders
        Result = StringCombine(vbCrLf, _
            Array(Result, SubFolder.Path, _
                FolderPathListSubFolder(SubFolder.Path)))
    Next
    FolderPathListSubFolder = Result
End Function

'----------------------------------------
'・ファイルリストを取得
'----------------------------------------
'   ・  存在しなければ空文字を返す。
'   ・  パスは改行コードで区切られている
'----------------------------------------
Public Function FilePathListTopFolder(ByVal FolderPath)
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FilePathListTopFolder:Folder no Exists")
    Dim Result: Result = ""
    Dim File
    For Each File In fso.GetFolder(FolderPath).Files
        Result = StringCombine(vbCrLf, Array(Result, File.Path))
    Next
    FilePathListTopFolder = ExcludeLastStr(Result, vbCrLf)
End Function

Public Function FilePathListSubFolder(ByVal FolderPath)
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FilePathListSubFolder:Folder no Exists")
    Dim Result: Result = ""
    Dim FolderPathList
    FolderPathList = Split( _
        FolderPathListSubFolder(FolderPath) + vbCrLf + FolderPath, vbCrLf)
    Dim I
    For I = 0 To ArrayCount(FolderPathList) - 1
        If fso.FolderExists(FolderPathList(I)) Then
            Result = StringCombine(vbCrLf, _
                Array(Result, FilePathListTopFolder(FolderPathList(I))))
        End If
    Next
    FilePathListSubFolder = ExcludeLastStr(Result, vbCrLf)
End Function


'----------------------------------------
'◆ファイルフォルダ操作
'----------------------------------------

'----------------------------------------
'・ファイルフォルダの名前変更
'----------------------------------------
Public Sub RenameFileFolder(ByVal Path, ByVal NewName)
    If fso.FileExists(Path) Then

        Call fso.MoveFile( _
            Path, _
            PathCombine(Array( _
                fso.GetParentFolderName(Path), _
                NewName _
            )) )

    ElseIf fso.FolderExists(Path) Then

        Call fso.MoveFolder( _
            Path, _
            PathCombine(Array( _
                fso.GetParentFolderName(Path), _
                NewName _
            )) )
    Else
        Call Assert(False, _
            "Error:RenameFileFolder:File/Folder not found.")
    End If
End Sub

'----------------------------------------
'・MoveFolder上書き可能
'----------------------------------------
'   ・  
'----------------------------------------
Sub MoveFolderOverWrite(ByVal SourceFolderPath, ByVal DestFolderPath)
    Call fso.CopyFolder(SourceFolderPath, DestFolderPath, True)
    Call ForceDeleteFolder(SourceFolderPath)
End Sub

'----------------------------------------
'◇Force/Recrate
'----------------------------------------

'----------------------------------------
'・深い階層のフォルダでも一気に作成する関数
'----------------------------------------
Public Sub ForceCreateFolder(ByVal FolderPath)
    Dim ParentFolderPath
    ParentFolderPath = fso.GetParentFolderName(FolderPath)

    If fso.FolderExists(ParentFolderPath) = False Then
        Call ForceCreateFolder(ParentFolderPath)
    End If

    'フォルダが出来るまで繰り返す。
    '100回繰り返して無理ならエラー
    Dim I: I = 1
    On Error Resume Next
    Do Until fso.FolderExists(FolderPath)
        Call fso.CreateFolder(FolderPath)
        If 100 <= I Then Exit Do
        WScript.Sleep 100
        I = I + 1
    Loop
    On Error GoTo 0
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:ForceCreateFolder:Create Folder Fail.")
End Sub

Private Sub testForceCreateFolder
    Call ForceDeleteFolder(AbsolutePath(ScriptFolderPath, _
        ".\Test\TestForceCreateFolder"))

    Call ForceCreateFolder(AbsolutePath(ScriptFolderPath, _
        ".\Test\TestForceCreateFolder\Test\Test"))

    Call Check(True, fso.FolderExists( AbsolutePath(ScriptFolderPath, _
        ".\Test\TestForceCreateFolder\Test\Test") ))

    Call ForceDeleteFolder(AbsolutePath(ScriptFolderPath, _
        ".\Test\TestForceCreateFolder"))
End Sub

'----------------------------------------
'・フォルダ削除を確認するまでDeleteFolderする関数
'----------------------------------------
Public Sub ForceDeleteFolder(ByVal FolderPath)
    'フォルダがある間繰り返す。
    '100回繰り返して無理ならエラー
    Dim I: I = 1
    On Error Resume Next
    Do While fso.FolderExists(FolderPath)
        Call fso.DeleteFolder(FolderPath)
        If 100 <= I Then Exit Do
        WScript.Sleep 100
        I = I + 1
    Loop
    On Error GoTo 0
    Call Assert(fso.FolderExists(FolderPath) = False, _
        "Error:ForceDeleteFolder:Folder Delete Fail.")
    'fso.DeleteFolderはファイルやサブフォルダあってもすべて消してくれる
End Sub


'----------------------------------------
'・ファイル削除を確認するまでDeleteFileする関数
'----------------------------------------
Public Sub ForceDeleteFile(ByVal FilePath)
    'ファイルがある間繰り返す。
    '100回繰り返して無理ならエラー
    Dim I: I = 1
    On Error Resume Next
    Do While fso.FileExists(FilePath)
        Call fso.DeleteFile(FilePath, True)
        If 100 <= I Then Exit Do
        WScript.Sleep 100
        I = I + 1
    Loop
    On Error GoTo 0
    Call Assert(fso.FileExists(FilePath) = False, _
        "Error:ForceDeleteFile:File Delete Fail.")
End Sub

'----------------------------------------
'・フォルダを再生成する関数
'----------------------------------------
Public Sub RecreateFolder(ByVal FolderPath)
    Call ForceDeleteFolder(FolderPath)
    Call ForceCreateFolder(FolderPath)
End Sub

'----------------------------------------
'・フォルダを再生成してコピーする関数
'----------------------------------------
'フォルダの日付が最新になる
'----------------------------------------
Public Sub RecreateCopyFolder( _
ByVal SourceFolderPath, ByVal DestFolderPath)

    Call Assert(fso.FolderExists(SourceFolderPath) = True, _
        "Error:ReCreateCopyFolder:Copy SourceFolder is not exists")

    Call ForceDeleteFolder(DestFolderPath)

    Dim I: I = 1
    On Error Resume Next
    Do
        Call ForceCreateFolder(fso.GetParentFolderName(DestFolderPath))
        Call fso.CopyFolder( _
            SourceFolderPath, DestFolderPath, True)
        If 100 <= I Then Exit Do
        WScript.Sleep 100
        I = I + 1
    Loop Until fso.FolderExists(DestFolderPath)
    'フォルダが作成できるまでループ
    On Error GoTo 0
    Call Assert(fso.FolderExists(DestFolderPath), _
        "Error:ReCreateCopyFolder:Copy Folder Fail.")
End Sub

'----------------------------------------
'◇除外ファイル/フォルダ指定
'----------------------------------------

'----------------------------------------
'・上書き除外ファイルを指定したフォルダ内ファイルのコピー関数
'----------------------------------------
'   ・  インストール時などにiniファイルを除外して
'       上書きインストールするときに使用する
'   ・  指定例：
'       IgnorePathsStr = "*.ini"
'       IgnorePathsStr = "*.ini,setting.txt"
'----------------------------------------
Public Sub CopyFolderOverWriteIgnorePath( _
ByVal SourceFolderPath, ByVal DestFolderPath, _
ByVal IgnorePathsStr)
    Call CopyFolderIgnorePath( _
        SourceFolderPath, DestFilePath, _
        "", IgnorePathsStr)
End Sub

'----------------------------------------
'・除外ファイル/フォルダを指定したフォルダ内ファイルのコピー関数
'----------------------------------------
'   ・  IgnorePathsStr はカンマ区切りで複数指定可能
'----------------------------------------
Public Sub CopyFolderIgnorePath( _
ByVal SourceFolderPath, ByVal DestFolderPath, _
ByVal IgnorePathsStr, ByVal OverWriteIgnorePathsStr)

    Dim FileList: FileList = _
        Split( _
            FilePathListSubFolder(SourceFolderPath), vbCrLf)

    Dim CopyDestFilePath
    Dim I
    For I = 0 To ArrayCount(FileList) - 1
    Do
        CopyDestFilePath = _
            IncludeFirstStr( _
                ExcludeFirstStr(FileList(I), SourceFolderPath), _
                DestFolderPath)

        Call CopyFileIgnorePathBase( _
            FileList(I), CopyDestFilePath, _
            IgnorePathsStr, OverWriteIgnorePathsStr)
    Loop While False
    Next
End Sub

'----------------------------------------
'・除外ファイル/フォルダを指定したファイルコピー関数
'----------------------------------------
'   ・  IgnorePathsStr はカンマ区切りで複数指定可能
'   ・  該当するファイルを普通に無視するCopyFileIgnorePathと
'       上書きだけ無視するCopyFileOverWriteIgnorePathがある
'----------------------------------------
Public Sub CopyFileIgnorePath( _
ByVal SourceFilePath, ByVal DestFilePath, _
ByVal IgnorePathsStr)
    Call CopyFileIgnorePathBase( _
        SourceFilePath, DestFilePath, IgnorePathsStr, "")
End Sub

Public Sub CopyFileOverWriteIgnorePath( _
ByVal SourceFilePath, ByVal DestFilePath, _
ByVal IgnorePathsStr)
    Call CopyFileIgnorePathBase( _
        SourceFilePath, DestFilePath, "", IgnorePathsStr)
End Sub

Private Sub CopyFileIgnorePathBase( _
ByVal SourceFilePath, ByVal DestFilePath, _
ByVal IgnorePathsStr, _
ByVal OverWriteIgnorePathsStr)

    Call Assert(fso.FileExists(SourceFilePath), _
        "Error:CopyFileIgnorePathBase:SourceFilePath is not exists")
    Do
        '除外ファイル
        If MatchText(LCase(SourceFilePath), _
            Split(LCase(IgnorePathsStr), ",")) Then
            Exit Do
        End IF

        '上書き除外ファイル
        If MatchText(LCase(SourceFilePath), _
            Split(LCase(OverWriteIgnorePathsStr), ",")) Then
            If fso.FileExists(DestFilePath) Then
                Exit Do
            End If
        End IF

        Call ForceCreateFolder(fso.GetParentFolderName(DestFilePath))
        Call fso.CopyFile( _
            SourceFilePath, DestFilePath, True)
    Loop While False
End Sub

'----------------------------------------
'・除外/対象パス指定ファイル削除処理
'----------------------------------------
'   ・  Target/IgnorePathsStr はカンマ区切りで複数指定可能
'----------------------------------------
Public Sub DeleteFileTargetPath( _
ByVal FileFolderPath, _
ByVal TargetPathsStr)
    If fso.FileExists(FileFolderPath) Then
        Call DeleteFileTargetIgnorePathBase( _
            FileFolderPath, TargetPathsStr, False)
    ElseIf fso.FolderExists(FileFolderPath) Then
        Dim FileList: FileList = _
            Split( _
                FilePathListSubFolder(FileFolderPath), vbCrLf)

        Dim DeleteFilePath
        Dim I
        For I = 0 To ArrayCount(FileList) - 1
        Do
            Call DeleteFileTargetIgnorePathBase( _
                FileList(I), TargetPathsStr, False)
        Loop While False
        Next
    End If
End Sub

Public Sub DeleteFileIgnorePath( _
ByVal FileFolderPath, _
ByVal IgnorePathsStr)
    If fso.FileExists(FileFolderPath) Then
        Call DeleteFileTargetIgnorePathBase( _
            FileFolderPath, IgnorePathsStr, True)
    ElseIf fso.FolderExists(FileFolderPath) Then
        Dim FileList: FileList = _
            Split( _
                FilePathListSubFolder(FileFolderPath), vbCrLf)

        Dim DeleteFilePath
        Dim I
        For I = 0 To ArrayCount(FileList) - 1
        Do
            Call DeleteFileTargetIgnorePathBase( _
                FileList(I), IgnorePathsStr, True)
        Loop While False
        Next
    End If
End Sub

Private Sub DeleteFileTargetIgnorePathBase( _
ByVal DeleteFilePath, _
ByVal PathsStr, _
ByVal DeleteIgnoreMode)
    Dim DeleteTargetMode: DeleteTargetMode = not DeleteIgnoreMode
    'DeleteIgnoreモードトDeleteTargetModeで処理を判断するため

    Call Assert(fso.FileExists(DeleteFilePath), _
        "Error:DeleteFileTargetIgnorePathBase:DeleteFilePath is not exists")
    Do
        '除外ファイル
        If MatchText(LCase(DeleteFilePath), _
            Split(LCase(PathsStr), ",")) Then
            If DeleteIgnoreMode Then
                Exit Do
            End If
        Else
            If DeleteTargetMode Then
                Exit Do
            End If
        End IF

        Call ForceDeleteFile(DeleteFilePath)
    Loop While False
End Sub

'----------------------------------------
'◇エラーを無視する処理
'----------------------------------------

'----------------------------------------
'・CopyFile
'----------------------------------------
'   ・  fso.CopyFileの最終要素に"*.*"を指定すると、
'       ファイルがない場合にエラーになるので
'       それを無視するための関数
'----------------------------------------
Sub CopyFile(ByVal SourceFilePath, ByVal DestFilePath)
On Error Resume Next
    Call fso.CopyFile(SourceFilePath, DestFilePath, True)
End Sub

'----------------------------------------
'・MoveFile
'----------------------------------------
'   ・  fso.MoveFileの最終要素に"*.*"を指定すると、
'       ファイルがない場合にエラーになるので
'       それを無視するための関数
'----------------------------------------
Sub MoveFile(ByVal SourceFilePath, ByVal DestFilePath)
On Error Resume Next
    Call fso.MoveFile(SourceFilePath, DestFilePath)
End Sub

'----------------------------------------
'・CopyFolder
'----------------------------------------
'   ・  fso.CopyFolderの最終要素に"*"を指定すると、
'       フォルダがない場合にエラーになるので
'       それを無視するための関数
'----------------------------------------
Sub CopyFolder(ByVal SourceFolderPath, ByVal DestFolderPath)
On Error Resume Next
    Call fso.CopyFolder(SourceFolderPath, DestFolderPath, True)
End Sub

'----------------------------------------
'・MoveFolder
'----------------------------------------
'   ・  fso.MoveFolderの最終要素に"*"を指定すると、
'       フォルダがない場合にエラーになるので
'       それを無視するための関数
'----------------------------------------
Sub MoveFolder(ByVal SourceFolderPath, ByVal DestFolderPath)
On Error Resume Next
    Call fso.MoveFolder(SourceFolderPath, DestFolderPath)
End Sub



'----------------------------------------
'◆ファイルの状態確認
'----------------------------------------

'----------------------------------------
'・ファイルフォルダ存在確認
'----------------------------------------
Public Function FileFolderExists(ByVal Path)
    Dim Result: Result = False
    If fso.FileExists(Path) _
    Or fso.FolderExists(Path) Then
        Result = True
    End If
    FileFolderExists = Result
End Function

'----------------------------------------
'・ファイル使用中かどうかを確認する
'----------------------------------------
Function IsUseFile(FilePath)
    Dim Result: Result = False
    Do
        If not fso.FileExists(FilePath) Then Exit Do
        If IsReadOnlyFile(FilePath) Then Exit Do
        On Error Resume Next
        Dim File
        Set File = fso.OpenTextFile(FilePath, 8, False)
        If 0 < Err.Number Then
            Result = True
        End If
    Loop While False
    IsUseFile = Result
End Function

'----------------------------------------
'・ファイルが読み取り専用かどうか確認する
'----------------------------------------
Function IsReadOnlyFile(FilePath)
    Dim Result: Result = False
    Do
        If fso.FileExists(FilePath) Then Exit Do
        '読み取り属性調査は定数1とのAndで判定する
        If fso.GetFile(FilePath).Attributes And 1 Then
            Result = True
        End If
    Loop While False
    IsReadOnlyFile = Result
End Function


'----------------------------------------
'◆ファイルフォルダ日時
'----------------------------------------

'----------------------------------------
'・ファイル/フォルダの最終更新日時を得る
'----------------------------------------
Public Function FileFolderDate_LastModified(ByVal Path)
    Dim Result: Result = Now()
    If fso.FileExists(Path) Then
        Result = fso.GetFile(Path).DateLastModified
    ElseIf fso.FolderExists(Path) Then
        Result = fso.GetFolder(Path).DateLastModified
    Else
        Call Assert(False, _
            "Error:FileFolderDate_LastModified:File/Folder not found.")
    End If
    FileFolderDate_LastModified = Result
End Function


'----------------------------------------
'◆ショートカットファイル操作
'----------------------------------------

'----------------------------------------
'・ショートカットファイルのリンク先を取得する
'----------------------------------------
Public Function ShortcutFileLinkPath(ByVal ShortcutFilePath)
    Dim Result: Result = ""
    Dim file: Set file = Shell.CreateShortcut(ShortcutFilePath)
    Result = file.TargetPath
    ShortcutFileLinkPath = Result
End Function


'----------------------------------------
'◆テキストファイル読み書き
'----------------------------------------

Public Function CheckEncodeName(ByVal EncodeName)
    CheckEncodeName = OrValue(UCase(EncodeName), Array(_
        "SHIFT_JIS", _
        "UNICODE", "UNICODEFFFE", "UTF-16LE", "UTF-16", _
        "UNICODEFEFF", _
        "UTF-16BE", _
        "UTF-8", _
        "UTF-8N", _
        "ISO-2022-JP", _
        "EUC-JP", _
        "UTF-7") )
End Function

'----------------------------------------
'・テキストファイル読込
'----------------------------------------
'   ・  エンコード指定は下記の通り
'           エンコード          指定文字
'           ShiftJIS            SHIFT_JIS
'           UTF-16LE BOM有/無   UNICODEFFFE/UNICODE/UTF-16/UTF-16LE
'                           BOMの有無に関わらず読込可能
'           UTF-16BE _BOM_ON    UNICODEFEFF
'           UTF-16BE _BOM_OFF   UTF-16BE
'           UTF-8 BOM有/無      UTF-8/UTF-8N
'                           BOMの有無に関わらず読込可能
'           JIS                 ISO-2022-JP
'           EUC-JP              EUC-JP
'           UTF-7               UTF-7
'   ・  UTF-16LEとUTF-8は、BOMの有無にかかわらず読み込める
'----------------------------------------
Const StreamTypeEnum_adTypeBinary = 1
Const StreamTypeEnum_adTypeText = 2
Const StreamReadEnum_adReadAll = -1
Const StreamReadEnum_adReadLine = -2
Const SaveOptionsEnum_adSaveCreateOverWrite = 2

Public Function LoadTextFile( _
ByVal TextFilePath, ByVal EncodeName)

    If CheckEncodeName(EncodeName) = False Then
        Call Assert(False, "Error:LoadTextFile")
    End If

    Dim Stream: Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = StreamTypeEnum_adTypeText

    Select Case UCase(EncodeName)
    Case "UTF-8N"
        Stream.Charset = "UTF-8"
    Case Else
        Stream.Charset = EncodeName
    End Select
    Call Stream.Open
    Call Stream.LoadFromFile(TextFilePath)
    LoadTextFile = Stream.ReadText
    Call Stream.Close
End Function

Private Sub testLoadTextFile()
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\SJIS_File.txt", "Shift_JIS"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-7_File.txt", "UTF-7"))

    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-8_File.txt", "UTF-8"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-8_File.txt", "UTF-8N"))

    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-8N_File.txt", "UTF-8N"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-8N_File.txt", "UTF-8"))

    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16BE_BOM_OFF_File.txt", "UTF-16BE"))

    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16BE_BOM_ON_File.txt", "UNICODEFEFF"))

    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_OFF_File.txt", "UTF-16"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_OFF_File.txt", "UTF-16LE"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_OFF_File.txt", "UNICODE"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_OFF_File.txt", "UNICODEFFFE"))

    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_ON1_File.txt", "UTF-16"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_ON1_File.txt", "UTF-16LE"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_ON1_File.txt", "UNICODE"))
    Call Check("123ABCあいうえお", LoadTextFile("Test\TestLoadTextFile\UTF-16LE_BOM_ON1_File.txt", "UNICODEFFFE"))
End Sub

'----------------------------------------
'・テキストファイル保存
'----------------------------------------
'   ・  エンコード指定は下記の通り
'           エンコード          指定文字
'           ShiftJIS            SHIFT_JIS
'           UTF-16LE _BOM_ON    UNICODEFFFE/UNICODE/UTF-16
'           UTF-16LE _BOM_OFF    UTF-16LE
'           UTF-16BE _BOM_ON    UNICODEFEFF
'           UTF-16BE _BOM_OFF    UTF-16BE
'           UTF-8 _BOM_ON       UTF-8
'           UTF-8 _BOM_OFF       UTF-8N
'           JIS                 ISO-2022-JP
'           EUC-JP              EUC-JP
'           UTF-7               UTF-7
'   ・  UTF-16LEとUTF-8はそのままだと_BOM_ONになるので
'       BON無し指定の場合は特殊処理をしている
'----------------------------------------
Public Sub SaveTextFile(ByVal Text, _
ByVal TextFilePath, ByVal EncodeName)
    If CheckEncodeName(EncodeName) = False Then
        Call Assert(False, "Error:SaveTextFile")
    End If

    Dim Stream: Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = StreamTypeEnum_adTypeText

    Select Case UCase(EncodeName)
    Case "UTF-8N"
        Stream.Charset = "UTF-8"
    Case Else
        Stream.Charset = EncodeName
    End Select
    Call Stream.Open
    Call Stream.WriteText(Text)
    Dim ByteData
    Select Case UCase(EncodeName)
    Case "UTF-16LE"
        Stream.Position = 0
        Stream.Type = StreamTypeEnum_adTypeBinary
        Stream.Position = 2
        ByteData = Stream.Read
        Stream.Position = 0
        Call Stream.Write(ByteData)
        Call Stream.SetEOS
    Case "UTF-8N"
        Stream.Position = 0
        Stream.Type = StreamTypeEnum_adTypeBinary
        Stream.Position = 3
        ByteData = Stream.Read
        Stream.Position = 0
        Call Stream.Write(ByteData)
        Call Stream.SetEOS
    End Select
    Call Stream.SaveToFile(TextFilePath, _
        SaveOptionsEnum_adSaveCreateOverWrite)
    Call Stream.Close
End Sub

Private Sub testSaveTextFile()
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\SJIS_File.txt", "Shift_JIS")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-7_File.txt", "UTF-7")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-8_File.txt", "UTF-8")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-8N_File.txt", "UTF-8N")

    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-16BE_BOM_OFF_File.txt", "UTF-16BE")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-16BE_BOM_ON_File.txt", "UNICODEFEFF")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-16LE_BOM_OFF_File.txt", "UTF-16LE")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-16LE_BOM_ON1_File.txt", "UNICODEFFFE")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-16LE_BOM_ON2_File.txt", "UNICODE")
    Call SaveTextFile("123ABCあいうえお", "Test\TestSaveTextFile\UTF-16LE_BOM_ON3_File.txt", "UTF-16")
End Sub

'----------------------------------------
'◆Iniファイル操作クラス
'----------------------------------------

Class IniFile
    Private IniDic
    Private Path
    Private UpdateFlag

    Private Sub Class_Initialize
        Set IniDic = WScript.CreateObject("Scripting.Dictionary")
        Path = ""
        UpdateFlag = False
    End Sub

    Private Sub Class_Terminate
        Set IniDic = Nothing
    End Sub

    Public Sub Initialize(ByVal IniFilePath)
        Call Assert(fso.FolderExists(fso.GetParentFolderName(IniFilePath)), _
            "Error:IniFile.Initialize:IniFile Folder not found")
        IniDic.RemoveAll

        If fso.FileExists(IniFilePath) Then
            Dim IniFileText
            IniFileText = LoadTextFile(IniFilePath, "SHIFT_JIS")

            Dim IniFileLines: IniFileLines = _
                Split(IniFileText, vbCrLf)
            Dim Section: Section = ""
            Dim Ident: Ident = ""
            Dim I
            For I = 0 To ArrayCount(IniFileLines) - 1
            Do
                Dim Line: Line = Trim(IniFileLines(I))
                If IsFirstStr(Line, "[") And IsLastStr(Line, "]") Then
                    Section = Line
                    Exit Do
                End If
                If Section = "" Then Exit Do
                Ident = FirstStrFirstDelim(Line, "=")
                If Ident <> "" Then
                    Call IniDic.Add(Section+Ident, LastStrFirstDelim(Line, "="))
                End If
            Loop While False
            Next
        End If
        Path = IniFilePath
        UpdateFlag = False
    End Sub

    Private Function DictionaryKey( _
    ByVal Section, ByVal Ident)
        DictionaryKey = IncludeLastStr(IncludeFirstStr(Section, "["), "]") + Ident
    End Function

    Public Function SectionIdentExists( _
    ByVal Section, ByVal Ident)
        SectionIdentExists = _
            IniDic.Exists(DictionaryKey(Section, Ident))
    End Function

    Public Function SectionIdentDelete( _
    ByVal Section, ByVal Ident)
        If IniDic.Exists(DictionaryKey(Section, Ident)) Then
            Call IniDic.Remove(DictionaryKey(Section, Ident))
        End If
    End Function

    Public Function ReadString( _
    ByVal Section, ByVal Ident, ByVal DefaultValue)
        Dim Result
        Dim Key: Key = DictionaryKey(Section, Ident)
        If IniDic.Exists(Key) Then
            Result = IniDic(Key)
            If Result = "" Then Result = DefaultValue
            '空文字設定ならデフォルト値を代入
        Else
            Result = DefaultValue
        End If
        ReadString = Result
    End Function

    Public Sub WriteString( _
    ByVal Section, ByVal Ident, ByVal Value)
        Dim Key: Key = DictionaryKey(Section, Ident)
        If IniDic.Exists(Key) Then
            IniDic(Key) = Value
        Else
            IniDic.Add Key, Value
        End If
        UpdateFlag = True
    End Sub

    Public Sub Update
        If UpdateFlag = False Then Exit Sub
        'WriteStringを実行したときだけ
        'Updateメソッドが動作する

        Dim IniSectionDic: Set IniSectionDic = WScript.CreateObject("Scripting.Dictionary")
        Dim Keys1: Keys1 = IniDic.Keys

        Dim Section: Section = ""
        Dim I
        For I = 0 To ArrayCount(Keys1) - 1
            Section = FirstStrFirstDelim(Keys1(I), "]") + "]"
            If IniSectionDic.Exists(Section) = False Then
                IniSectionDic.Add Section, ""
            End If
        Next

        Dim IniFileText: IniFileText = ""
        Dim Keys2: Keys2 = IniSectionDic.Keys
        Dim J
        For J = 0 To ArrayCount(Keys2) - 1
            IniFileText = IniFileText + _
                Keys2(J) + vbCrLf
            For I = 0 To ArrayCount(Keys1) - 1
                If IsFirstStr(Keys1(I), Keys2(J)) Then
                    IniFileText = IniFileText + _
                        ExcludeFirstStr(Keys1(I), Keys2(J)) + _
                        "=" + IniDic(Keys1(I)) + vbCrLf
                End If
            Next
        Next
        Set IniSectionDic = Nothing

        Call SaveTextFile(IniFileText, Path, "SHIFT_JIS")
    End Sub
End Class

Sub testIniFile
    Dim IniFilePath: IniFilePath = _
        ScriptFolderPath + "\Test\TestIniFile\" + _
        "testIniFile.ini"

    Call ForceCreateFolder(fso.GetParentFolderName(IniFilePath))

    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)

    Call IniFile.WriteString( _
        "Option1", "Test01", "ValueA")
    Call IniFile.WriteString( _
        "Option2", "Test01", "ValueB")
    Call IniFile.WriteString( _
        "Option1", "Test01", "ValueC")

    IniFile.Update

    Call Check("ValueC", IniFile.ReadString("Option1", "Test01", ""))
    Call Check("ValueB", IniFile.ReadString("Option2", "Test01", ""))

    Set IniFile = Nothing
End Sub


'----------------------------------------
'◆圧縮ファイル処理
'----------------------------------------

'----------------------------------------
'・Zipファイル解凍
'----------------------------------------
'   ・  Windows標準機能を利用
'   ・  テンポラリフォルダに展開してから移動する
'----------------------------------------
Function UnZip(ByVal ZipFilePath, ByVal UnCompressFolderPath)
    Call Assert(fso.FileExists(ZipFilePath), "Error:UnZip:ZipFilePath no exists")
    Call Assert(fso.FolderExists(UnCompressFolderPath), "Error:UnZip:UnCompressFolderPath no exists")

    Dim OutputTempFolderPath: OutputTempFolderPath = _
        PathCombine(Array(UnCompressFolderPath, TemporaryFolderName))
    Call RecreateFolder(OutputTempFolderPath)

    Dim Shell_Application
    Set Shell_Application = WScript.CreateObject("Shell.Application")
    Call Shell_Application.Namespace(OutputTempFolderPath).CopyHere( _
        Shell_Application.Namespace(ZipFilePath).Items)

    Call CopyFile(PathCombine(Array(OutputTempFolderPath, "*.*")), _
        IncludeLastPathDelim(UnCompressFolderPath))
    Call CopyFolder(PathCombine(Array(OutputTempFolderPath, "*")), _
        IncludeLastPathDelim(UnCompressFolderPath))

    Call ForceDeleteFolder(OutputTempFolderPath)
End Function

'----------------------------------------
'・Zipファイル圧縮
'----------------------------------------
'   ・  Windows標準機能を利用
'   ・  テンポラリファイルを作成してから名前変更している
'----------------------------------------
Function Zip(ByVal CompressSourcePath, ByVal ZipFilePath)
    Call Assert(fso.FileExists(ZipFilePath) = False, "Error:Zip:ZipFilePath exists")
    Call Assert(FileFolderExists(CompressSourcePath), "Error:Zip:CompressSourcePath no exists")

    Dim OutputTempZipFilePath
    Do
        OutputTempZipFilePath = _
            PathCombine(Array( _
                fso.GetParentFolderName(ZipFilePath), TemporaryFileName))
        OutputTempZipFilePath = ChangeFileExt(OutputTempZipFilePath, ".zip")
    Loop While fso.FileExists(OutputTempZipFilePath)
    '一時ファイルのパスを指定する
    '拡張子はzipでないと動作しないので注意

    Dim Data: Data = _
        Chr(80) +  Chr(75) + Chr(5) + Chr(6) + _
        String(18, Chr(0))
    '空のZIPファイルの作成
    Call SaveTextFile(Data, OutputTempZipFilePath, "SHIFT_JIS")

    Dim Shell_Application
    Set Shell_Application = WScript.CreateObject("Shell.Application")
    Call Shell_Application.Namespace(OutputTempZipFilePath).CopyHere( _
        CompressSourcePath)

    'ファイルができるまで待機
    Do: Call WScript.Sleep(100)
    Loop Until fso.FileExists(OutputTempZipFilePath)

    'ファイルが使用状態の間待機
    Do: Call WScript.Sleep(100)
    Loop While IsUseFile(OutputTempZipFilePath)

    Call RenameFileFolder(OutputTempZipFilePath, fso.GetFileName(ZipFilePath))
End Function

'----------------------------------------
'◆システム
'----------------------------------------

'----------------------------------------
'・実行時GUI/CUI確認
'----------------------------------------

Function IsCUI
    IsCui = _
        IsFirstStr(LCase(WScript.FullName), "cscript.exe")
End Function

Function IsGUI
    IsCui = _
        IsFirstStr(LCase(WScript.FullName), "wscript.exe")
End Function

'----------------------------------------
'◆コマンドライン引数
'----------------------------------------

'----------------------------------------
'・コマンドライン引数を配列に入れる関数
'----------------------------------------
Public Function ArgsToArray
    Dim Result()
    Dim Args: Set Args = WScript.Arguments
    ReDim Result(Args.Count - 1)
    Dim I
    For I = 0 To Args.Count - 1
        Result(I) = Args(I)
    Next
    ArgsToArray = Result
End Function

'----------------------------------------
'・引数からオプションを調べる
'----------------------------------------
'   ・  OptionFirstStrArray = Array("/a", "-a") のように指定する
'   ・  そのオプションで開始する引数が存在すればその引数を返す。
'       存在しなければ空文字を返す
'       例：[/a:123 /b /c file.txt]というコマンドラインから
'       ArgsOption(ArgsArray, Array("/a"))を指定すると
'       戻り値は "/a:123" になる
'   ・  該当オプションが2つ存在しても先頭のものだけ取り出す
'----------------------------------------
Public Function ArgsOption(ByRef ArgsArray, ByVal OptionFirstStrArray)
    Call Assert(IsArray(ArgsArray), _
        "Error:Function ArgsOption:ArgsArray is not array")
    Call Assert(IsArray(OptionFirstStrArray), _
        "Error:ArgsOptionCheckError:OptionFirstStrArray is not array")

    Dim OptionStr: OptionStr = ""
    Dim I
    For I = UBound(ArgsArray) To 0 Step -1
        Dim J
        For J = LBound(OptionFirstStrArray) To UBound(OptionFirstStrArray)
            If IsFirstStr(LCase(ArgsArray(I)), LCase(OptionFirstStrArray(J))) Then
                OptionStr = ArgsArray(I)
                Call ArrayDelete(ArgsArray, I)
                Exit For
            End If
        Next
    Next
    ArgsOption = OptionStr
End Function

Private Sub testArgsOption
    Dim Args: Args = Array("/a:123", "/b", "/a:456", "/c", "file.txt")
    Call Check("/a:123", ArgsOption(Args, Array("/a")))
    Call Check("/c", ArgsOption(Args, Array("/c")))
    Call Check("", ArgsOption(Args, Array("/d")))
End Sub

'----------------------------------------
'・引数からオプションのエラーを調べる
'----------------------------------------
'   ・  OptionFirstStrArray = Array("/a", "-a") のように指定する
'   ・  重複オプションがあれば True を返す
'       例：[/a:123 -a /c file.txt]というコマンドラインから
'       ArgsOption(ArgsArray, Array("/a", "-a"))を指定すると
'       該当オプションが2つあるのでエラーということでTrueを返す
'----------------------------------------
'引数からオプションを調べる
'同じものが二つみつかれば True = エラーがあるということ
Public Function ArgsOptionCheckError(ByRef ArgsArray, ByVal OptionFirstStrArray)
    Call Assert(IsArray(ArgsArray), _
        "Error:ArgsOptionCheckError:ArgsArray is not array")
    Call Assert(IsArray(OptionFirstStrArray), _
        "Error:ArgsOptionCheckError:OptionFirstStrArray is not array")

    Dim Result: Result = False
    Dim OptionStr: OptionStr = ""
    Dim I
    For I = UBound(ArgsArray) To 0 Step -1
        Dim J
        For J = LBound(OptionFirstStrArray) To UBound(OptionFirstStrArray)
            If IsFirstStr(LCase(ArgsArray(I)), LCase(OptionFirstStrArray(J))) Then
                If OptionStr <> "" Then
                    Result = True
                End If
                OptionStr = ArgsArray(I)
            End If
        Next
    Next
    ArgsOptionCheckError = Result
End Function

Private Sub testArgsOptionCheckError
    Dim Args: Args = Array("/a:123", "/b", "/a:456", "-b", "file.txt")
    Call Check(True, ArgsOptionCheckError(Args, Array("/a")))
    Call Check(False, ArgsOptionCheckError(Args, Array("/b")))
    Call Check(True, ArgsOptionCheckError(Args, Array("/b", "-b")))
    Call Check(False, ArgsOptionCheckError(Args, Array("/c")))
End Sub


'----------------------------------------
'◆クリップボード
'----------------------------------------

'----------------------------------------
'・テキストデータ取得
'----------------------------------------
Public Function GetClipboardText
    Dim HtmlFile
    Set HtmlFile = WScript.CreateObject("HtmlFile")
    GetClipboardText = HtmlFile.ParentWindow.ClipboardData.GetData("text")
End Function

'----------------------------------------
'・テキストデータ設定
'----------------------------------------
'   ・  IEオブジェクト利用のクリップボード設定方法などは
'       セキュリティの関係で安定して動作しないようなので
'       一時ファイルとClipコマンドを使用した
'----------------------------------------
Public Sub SetClipboardText(ByVal ClipboardToText)
    Dim TempFileName: TempFileName = TemporaryPath

    Call SaveTextFile(ClipboardToText, TempFileName, "Shift_JIS")

    Call Shell.Run( _
        "%ComSpec% /c clip < " + _
        InSpacePlusDoubleQuote(TempFileName), _
        0, True)

    Call ForceDeleteFile(TempFileName)
End Sub

'ExcelなどOfficeのある環境の場合は次のコードでも可能
'Sub SetClipboardText(ByVal Value)
'    Dim FormObj: Set FormObj = CreateObject("Forms.Form.1")
'    Dim TextBoxObj: Set TextBoxObj=FormObj.Controls.Add("Forms.TextBox.1").Object
'    TextBoxObj.MultiLine=True
'    TextBoxObj.Text=Value
'    TextBoxObj.SelStart=0
'    TextBoxObj.SelLength = TextBoxObj.TextLength
'    TextBoxObj.Copy
'End Sub

'----------------------------------------
'◆シェル
'----------------------------------------

'----------------------------------------
'・シェル起動定数
'----------------------------------------
Const vbHide = 0             'ウィンドウ非表示
Const vbNormalFocus = 1      '通常表示起動
Const vbMinimizedFocus = 2   '最小化起動
Const vbMaximizedFocus = 3   '最大化起動
Const vbNormalNoFocus = 4    '通常表示起動、フォーカスなし
Const vbMinimizedNoFocus = 6 '最小化起動、フォーカスなし

'----------------------------------------
'・ファイル指定したシェル起動
'----------------------------------------
Public Sub ShellFileOpen( _
ByVal FilePath, ByVal Focus)
    Call Assert(OrValue(Focus, Array(0, 1, 2, 3, 4, 6)), "Error:ShellFileOpen")

    Call Shell.Run( _
        "rundll32.exe url.dll" & _
        ",FileProtocolHandler " + FilePath _
        , Focus, False)
    'ファイル起動の場合
    '第三引数のWaitはTrueにしても無視される様子
End Sub

Private Sub testShellFileOpen
    Dim Path: Path = PathCombine(Array( _
            ScriptFolderPath, _
            "Test\TestShellRun\TestShellRun.txt"))
    Call ShellFileOpen(Path, vbNormalFocus)
    Call ShellFileOpen(Path, vbNormalFocus)
End Sub

'----------------------------------------
'・コマンド指定したシェル起動
'----------------------------------------
'   ・  Wait = Trueならプログラムの終了を待つ
'          Falseならそのまま実行を続ける
'----------------------------------------
Public Sub ShellCommandRun(Command, Focus, Wait)
    Call Assert(OrValue(Focus, Array(0, 1, 2, 3, 4, 6)), "Error:ShellCommandRun")
    Call Assert(OrValue(Wait, Array(True, False)), "Error:ShellCommandRun")

    Call Shell.Run( _
        Command _
        , Focus, Wait)
End Sub

Private Sub testShellCommandRun
    Dim Path: Path = "notepad.exe"
    Call ShellCommandRun(Path, vbNormalFocus, True)
    Call ShellCommandRun(Path, vbNormalFocus, False)
    Call ShellCommandRun(Path, vbNormalFocus, True)
End Sub

'----------------------------------------
'・DOSコマンドの実行。戻り値取得。
'----------------------------------------
'   ・  DOS窓が表示されない
'----------------------------------------
Public Function ShellCommandRunReturn(Command, Focus)
    Call Assert(OrValue(Focus, Array(0, 1, 2, 3, 4, 6)), "Error:ShellCommandRun")

    Dim TempFileName: TempFileName = TemporaryPath

    Call Shell.Run( _
        "%ComSpec% /c " + Command + ">" + TempFileName + " 2>&1" _
               , Focus, True)
    ' 戻り値を取得
    ShellCommandRunReturn = _
        LoadTextFile(TempFileName, "shift_jis")

    Call ForceDeleteFile(TempFileName)
End Function

Sub testShellCommandRunReturn()
    MsgBox ShellCommandRunReturn("tree ""C:\Software""", vbHide)
    MsgBox ShellCommandRunReturn("dir C:\", vbHide)
End Sub

'----------------------------------------
'・環境変数の取得
'----------------------------------------
Public Function EnvironmentalVariables(ByVal Name)
On Error Resume Next
    Dim Result: Result = ""
    Result = Shell.ExpandEnvironmentStrings( _
        IncludeBothEndsStr(Name, "%"))
    EnvironmentalVariables = Result
End Function

Private Sub testEnvironmentalVariables
    MsgBox EnvironmentalVariables("%TEMP%")
    MsgBox EnvironmentalVariables("%username%")
    MsgBox EnvironmentalVariables("appdata")
End Sub

'----------------------------------------
'◆プロセス存在チェック
'----------------------------------------
' プロセスが起動しているか調べる
'----------------------------------------
Public Function ProcessExists1(ByVal ProcessName)
    Dim Result: Result = False
    Dim Service: Set Service = _
        WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Dim DataSet: Set DataSet = Service.ExecQuery( _
        "SELECT * FROM Win32_Process")
    Dim Data
    For Each Data In DataSet
        If Data.Description = ProcessName Then
            Result = True
        End If
    Next
    ProcessExists1 = Result
End Function

Public Function ProcessExists2(ByVal ProcessName)
    Dim Result: Result = False
    Dim Service: Set Service = _
        WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Dim DataSet: Set DataSet = Service.ExecQuery( _
        "SELECT * FROM Win32_Process WHERE Name='" & ProcessName & "'")
    Dim Data
    For Each Data In DataSet
        Result = True
    Next
    ProcessExists2 = Result
End Function

Public Function ProcessExists(ByVal ProcessName)
    Dim Query: Query = _
        "SELECT * FROM Win32_Process WHERE Name = '" + ProcessName + "'"
    Dim Service: Set Service = GetObject("winmgmts:")
    Dim DataSet: Set DataSet = Service.ExecQuery(Query)
    ProcessExists = (DataSet.Count > 0)
End Function

Private Sub testProcessExists
    Call Check(True, ProcessExists("explorer.exe"))
    Call Check(True, ProcessExists1("explorer.exe"))
    Call Check(True, ProcessExists2("explorer.exe"))

    Call Check(False, ProcessExists("abc.exe"))
    Call Check(False, ProcessExists1("abc.exe"))
    Call Check(False, ProcessExists2("abc.exe"))

    'Call Check(False, ProcessExists("C:\Windows\explorer.exe"))
    'Call Check(False, ProcessExists1("C:\Windows\explorer.exe"))
    'Call Check(False, ProcessExists2("C:\Windows\explorer.exe"))
    '上記はエラーになる
End Sub


'----------------------------------------
'◆キーコード送信
'----------------------------------------

'----------------------------------------
'・ウィンドウタイトルを指定してキー送信
'----------------------------------------
'   ・  Shell.AppActivateを実行して成功してから
'       Shell.SendKeysを送信する関数
'   ・  SearchWindowTitle=""と指定すると
'       Shell.SendKeysだけの処理になる
'   ・  キーの文字は小文字で指定すること。
'       Ctrl+Cキーを指定しようとして[^C]と
'       大文字で指定するとShiftがロックされて挙動がおかしくなる
'----------------------------------------
Public Function AppActSendKeysLoop( _
ByVal SearchWindowTitle, _
ByVal KeyValue, ByVal WaitMilliSec, ByVal LoopCount)

    Dim Result: Result = False
    If SearchWindowTitle = "" Then
        Shell.SendKeys KeyValue
        WScript.Sleep WaitMilliSec
        Result = True
    Else
        Dim I
        For I = 1 To LoopCount
            If Shell.AppActivate(SearchWindowTitle, True) Then
                WScript.Sleep 100
                Shell.SendKeys KeyValue
                WScript.Sleep WaitMilliSec
                Result = True
                Exit For
            End If
            WScript.Sleep WaitMilliSec
        Next
    End If
    AppActSendKeysLoop = Result
End Function

Public Function AppActSendKeys( _
ByVal SearchWindowTitle, _
ByVal KeyValue, ByVal WaitMilliSec)
    AppActSendKeys = AppActSendKeysLoop( _
        SearchWindowTitle, _
        KeyValue, WaitMilliSec, 10)
End Function

'----------------------------------------
'・キー送信後、指定ウィンドウタイトルのアクティブ化確認
'----------------------------------------
'   ・  AppActSendKeysLoop後
'       配列で指定したウィンドウタイトルのどれかをアクティブに
'       出来たらTrueを返す関数
'   ・  例：    Call AppActSendKeysAfterWindow("test", _
'               Array("WindowA", "WindowB"), "%te", 1000)
'----------------------------------------
Public Function AppActSendKeysAfterWindowLoop( _
ByVal SearchWindowTitle, ByVal AfterWindowTitle, _
ByVal KeyValue, ByVal WaitMilliSec, ByVal LoopCount)
    Call Assert(IsArray(AfterWindowTitle), _
        "Error:AppActSendKeysAfterWindow:AfterWindowTitle is not Array.")

    Dim Result: Result = False
    Dim I: Dim J

    If AppActSendKeysLoop(SearchWindowTitle, KeyValue, WaitMilliSec, LoopCount) Then
        For J = 0 to ArrayCount(AfterWindowTitle) - 1
            If Shell.AppActivate(AfterWindowTitle(J), True) Then
                Result = True
                Exit For
            End If
        Next
    End If

    AppActSendKeysAfterWindowLoop = Result
End Function

Public Function AppActSendKeysAfterWindow( _
ByVal SearchWindowTitle, ByVal AfterWindowTitle, _
ByVal KeyValue, ByVal WaitMilliSec)
    AppActSendKeysAfterWindow = _
        AppActSendKeysAfterWindowLoop(_
        SearchWindowTitle, AfterWindowTitle, _
        KeyValue, WaitMilliSec, 10)
End Function


'----------------------------------------
'◆VBScript ソースコード編集
'----------------------------------------

'----------------------------------------
'・ソース中の[Call Include("ライブラリパス")]という記述を
'  ライブラリファイルの中身全てで置き換える処理
'----------------------------------------
Public Function IncludeExpanded( _
ByVal SourceFilePath, ByVal DestFilePath)

    Dim Result: Result = False
    Do
        If not fso.FileExists( _
            SourceFilePath) Then Exit Do

        If not fso.FolderExists( _
            fso.GetParentFolderName(DestFilePath)) Then Exit Do

        Call fso.CopyFile( _
            SourceFilePath, _
            DestFilePath, True)

        Dim FileText: FileText = _
            LoadTextFile(DestFilePath, "SHIFT_JIS")
        With Nothing
            Dim Lines: Lines = _
                Split(FileText, vbCrLf)
            Dim I
            For I = 0 To ArrayCount(Lines) - 1
                If 1 = InStr(Trim(Lines(I)), "Call Include(""") Then
                    Dim LibPath: LibPath = Trim(Lines(I))
                    LibPath = ExcludeFirstStr(LibPath, "Call Include(""")
                    LibPath = ExcludeLastStr(LibPath, """)")
                    LibPath = AbsolutePath( _
                        fso.GetParentFolderName(SourceFilePath), _
                        LibPath)
                    Call Assert(fso.FileExists(LibPath), _
                        "Error:IncludeExpanded:Inclde Library is not found.")
                    Lines(I) = _
                        Replace( _
                            LoadTextFile(LibPath, "SHIFT_JIS"), _
                            "Option Explicit", "")
                End If
            Next
            FileText = ArrayToString(Lines, vbCrLf)
        End With
        Call SaveTextFile(FileText, DestFilePath, "SHIFT_JIS")
    Loop While False
    IncludeExpanded = Result
End Function

'----------------------------------------
'・スクリプトエンコーダーを使ったVBSのエンコード
'----------------------------------------
Sub EncodeVBScriptFile(ByVal ScriptEncoderExePath, _
ByVal SourceFilePath, ByVal DestFilePath)
    Call ShellCommandRun( _
        InSpacePlusDoubleQuote(ScriptEncoderExePath) + _
        " " + InSpacePlusDoubleQuote(SourceFilePath) + _
        " " + InSpacePlusDoubleQuote(DestFilePath) _
        , vbHide, True)
End Sub

'--------------------------------------------------
'■履歴
'◇ ver 2015/01/20
'・ 作成
'・ Assert/Check/OrValue
'・ fso
'・ TestStandardSoftwareLibrary.vbs/Include
'・ LoadTextFile/SaveTextFile
'◇ ver 2015/01/21
'・ FirstStrFirstDelim/FirstStrLastDelim
'   /LastStrFirstDelim/LastStrLastDelim
'◇ ver 2015/01/26
'・ IsFirstStr/IncludeFirstStr/ExcludeFirstStr
'   /IsLastStr/IncludeLastStr/ExcludeLastStr
'◇ ver 2015/01/27
'・ ArrayCount
'・ TrimFirstStrs/TrimLastStrs/TrimBothEndsStrs
'・ StringCombine
'◇ ver 2015/02/02
'・ Shellオブジェクト
'・ CurrentDirectory/ScriptFolderPath
'・ AbsoluteFilePath
'   /IsNetworkPath/GetDrivePath
'   /IsIncludeDrivePath/IncludeLastPathDelim/ExcludeLastPathDelim
'   /PeriodExtName/ExcludePathExt/ChangeFileExt
'   /PathCombine
'・ IIF
'◇ ver 2015/02/04
'・ StringCombine修正
'・ FolderPathListTopFolder/FolderPathListSubFolder
'   /FilePathListTopFolder/FilePathListSubFolder
'・ ForceCreateFolder/ReCreateCopyFolder
'・ ShellFileOpen/ShellCommandRun/ShellCommandRunReturn
'・ EnvironmentalVariables
'・ FormatYYYY_MM_DD/FormatHH_MM_SS
'◇ ver 2015/02/05
'・ ForceDeleteFolder/RecreateFolder
'・ IniFile読み書きクラス作成
'・ IsCUI/IsGUI
'◇ ver 2015/02/06
'・ ArrayText
'・ LikeCompare/MatchText
'・ ShortcutFileLinkPath
'・ ForceCreateFolder修正
'◇ ver 2015/02/08
'・ ForceCreateFolder/ForceDeleteFolder修正
'・ CopyFolderOverWriteIgnoreFile作成
'◇ ver 2015/02/09
'・ IniFile処理をリファクタリング
'・ MaxValue/MinValue作成
'・ IsLong/LongToStrDigitZero/StrToLongDefault追加
'・ AppActSendKeys/AppActSendKeysAfterWindow等を追加
'◇ ver 2015/02/10
'・ InSpacePlusDoubleQuote修正
'・ ShellCommandRunReturnリファクタリング
'   TemporaryFilePathの作成
'・ GetClipbordText/SetClipbordText作成
'・ ForceDeleteFile作成
'◇ ver 2015/02/12
'・ FilePathListSubFolderの修正
'◇ ver 2015/02/16
'・ IncludeStr追加
'・ MatchTextWildCard/MatchTextKeyWord追加
'・ CopyFolderIgnoreFileFolder追加
'・ ArrayText廃止
'   IncludeBothEndsDoubleQuoteCombineArray追加
'◇ ver 2015/02/17
'・ LikeCompare修正
'・ TemporaryFilePath>>TemporaryPath名前変更
'◇ ver 2015/02/18
'・ TemporaryFolderPath追加
'◇ ver 2015/02/19
'・ ArrayAdd/ArrayDelete/ArrayInsert 追加
'・ CopyFileIgnoreFileFolderを追加
'   CopyFolderIgnoreFileFolderを修正
'・ UnZip機能追加
'・ CopyFile/CopyFolder追加
'◇ ver 2015/02/23
'・ TemporaryFilePath/TemporaryFolderName追加
'・ ForceDeleteFolder/ForceDeleteFile 
'   /ForceCreateFolder/ReCreateCopyFolder
'   例外発生時の不具合修正
'◇ ver 2015/02/26
'・ CopyFolderIgnoreFileFolderをCopyFolderIgnorePathに名称変更
'・ CopyFileIgnoreFileFolderをCopyFileIgnorePathに名称変更
'・ CopyFileIgnorePathBaseを作成して
'   CopyFileIgnorePath/CopyFileOverWriteIgnorePathと
'   処理を共通化
'・ CopyFolderOverWriteIgnoreをCopyFolderOverWriteIgnorePathに名称変更
'   内部をCopyFileOverWriteIgnorePathとして作成
'・ DeleteFileTargetPath/DeleteFileIgnorePathを追加
'◇ ver 2015/02/27
'・ IniFile.ReadStringの空項目取得時のデフォルト値処理を修正
'◇ ver 2015/03/02
'・ IsUseFile/IsReadOnlyFile追加
'・ ProcessExists追加
'・ IncludeExpanded/EncodeVBScriptFile追加
'◇ ver 2015/03/03
'・ ForceCreateFolder/ForceDeleteFolder
'   /ForceDeleteFile/ReCreateCopyFolder の待機時間を追加
'◇ ver 2015/03/06
'・ MoveFile/MoveFolderを追加
'・ DeleteDateTextを追加
'・ RenameFileFolder追加
'・ FileFolderDate_LastModified追加
'・ FileFolderExists追加
'◇ ver 2015/03/07
'・ MatchText/MatchTextWildCard/MatchTextKeyWordを
'   リファクタリングしてIndexOfArrayを作成
'・ MatchTextRegExpを追加
'◇ ver 2015/03/09
'・ IniFileクラスにSectionIdentDelete追加
'◇ ver 2015/03/17
'・ コメントや並び順などの変更
'・ TemporaryFileName作成
'・ Zip作成
'◇ ver 2015/03/18
'・ コメントの修正
'・ InRange追加
'・ SetValue追加
'・ ArrayAdd/ArrayDelete/ArrayInsertをオブジェクト値にも対応した
'・ ArgsToArray/ArgsOption/ArgsOptionCheckError を追加
'・ IsFirstText/IncludeFirstText/ExcludeFirstText
'   IsLastText/IncludeLastText/ExcludeLastText
'   IncludeBothEndsText/ExcludeBothEndsText を追加
'・ IncludeStr>>IsIncludeStrに名前変更
'◇ ver 2015/04/10
'・ DeleteDateTimeText追加
'・ コメント形式の全体的な修正
'◇ ver 2015/04/25
'・ コメント形式を変更
'◇ ver 2015/05/17
'・ CopyFolderIgnorePath内のファイルパスの大小文字不一致の不具合を修正
'・ StrToBooleanを追加
'・ FormatDateTimeToStringを追加
'◇ ver 2015/05/20
'・ DeleteDateText/DeleteDateTimeText を修正
'◇ ver 2015/07/23
'・ StarndardSoftwareLibraryからst.vbsに名前変更
'・ ライセンスをGPLからMITに変更した。
'・ MoveFolderOverWriteを追加
'◇ ver 2015/07/24
'・ CopyFileIgnorePathBase/CopyFolderIgnorePathを変更
'   通常無視と上書き無視を同時指定可能にした
'--------------------------------------------------
