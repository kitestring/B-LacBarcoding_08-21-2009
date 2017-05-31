Attribute VB_Name = "Module2"
'Function Module

Public Function SequentialSampleNumbers(ByVal RangeLetter As String, _
    ByVal ToggleButtonCaption As String) As String
        Sheets("Controls").Select
        Sheets("Controls").Unprotect Password:="12345"
        If ToggleButtonCaption = "Yes" Then
            Range(RangeLetter & "7:" & RangeLetter & "23").Select
            Selection.Locked = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            SequentialSampleNumbers = "No"
        ElseIf ToggleButtonCaption = "No" Then
            Range("A7").Copy
            Range(RangeLetter & "7:" & RangeLetter & "23").Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            CutCopyMode = False
            Range(RangeLetter & "7:" & RangeLetter & "23").Select
            Selection.Locked = True
            SequentialSampleNumbers = "Yes"
        End If
        Range(RangeLetter & "6").Select
        Sheets("Controls").Protect Password:="12345"
End Function

Public Function SampleInputsCheck() As Boolean
        Sheets("Controls").Select
        If Range("B1") = 0 Then
            SampleInputsCheck = True
        Else
            SampleInputsCheck = False
        End If
End Function

Public Function OpenDataBase(ByVal DB_NetAddress As Variant)
    Workbooks.OpenText Filename:=DB_NetAddress _
        , Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
        , Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
End Function

Public Function Location(ByVal Direction As String) As Integer
    If Direction = "Column" Then
        Selection.End(xlToRight).Select
        ActiveCell.Offset(0, 1).Select
        ActiveCell = "=(COLUMN())-1"
        Location = ActiveCell
        ActiveCell.ClearContents
        Range("A1").Select
    ElseIf Direction = "Row" Then
        ActiveCell.Offset(3, -1).Select
        ActiveCell = "=ROW()"
        Location = ActiveCell
        ActiveCell.ClearContents
        Range("A1").Select
    End If
End Function

Public Function FindStringInDataBase(ByVal strSearchFor As String, _
    Optional ByVal strLabel As String = "Nothing", Optional ByVal bolRight As Boolean = False) As Integer
        Cells.Find(What:=strSearchFor, After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        If bolRight = True Then
            FindStringInDataBase = Right(ActiveCell, 4)
        End If
        If Not (strLabel = "Nothing") Then
        End If
End Function

Public Function HighlightRawDataSlice(ByVal intSampleNumber As Integer, ByVal intStartRow As Integer, _
    ByVal intEndRow As Integer) As Boolean
        Dim i As Integer
        Dim a As Byte
        a = 0
        i = 0
        Do While a = 0
            If Cells(intStartRow + i, 2).Value = intSampleNumber Then
                Range(Cells(intStartRow + i, 2), Cells(intStartRow + i, intEndRow)).Select
                a = 1
                HighlightRawDataSlice = True
            ElseIf IsEmpty(Cells(intStartRow + i, 2)) Then
                a = 1
                HighlightRawDataSlice = False
            End If
            i = i + 1
        Loop
End Function
