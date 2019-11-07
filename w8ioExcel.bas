Sub SummOrdersForAds()

'save main address
mainA = Cells(1, 1).Value
sep = InStr(1, mainA, "/", 1)
If sep > 0 Then
    mainA = Left(mainA, sep - 2)
    'Debug.Print mainA
Else
    mainA = Left(mainA, InStr(1, mainA, " ", 1) - 1)
    'Debug.Print mainA
End If

'delete unwanted
Columns("A:A").Select
Selection.Delete Shift:=xlToLeft

Rows("1:2").Select
Selection.Delete Shift:=xlUp

'fix names of tx type
Cells.Replace What:="(smart ", Replacement:="(smart_", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
Cells.Replace What:="(invoke ", Replacement:="(invoke_", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
    :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
    Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1)), _
    TrailingMinusNumbers:=True

Columns("I:L").Select
Selection.Delete Shift:=xlToLeft

Columns("A:H").Select
Selection.Columns.AutoFit

With Application
    .DecimalSeparator = "."
    .UseSystemSeparators = False
End With

Columns("H:H").Select
Selection.NumberFormat = "0.00000000"

'delete all tx except input
txtype = InputBox("Enter tx type to analise" _
& vbNewLine _
& vbNewLine & "0 = (fees)" _
& vbNewLine & "1 = (genesis)" _
& vbNewLine & "2 = (payment)" _
& vbNewLine & "3 = (issue)" _
& vbNewLine & "4 = (transfer)" _
& vbNewLine & "5 = (reissue)" _
& vbNewLine & "6 = (burn)" _
& vbNewLine & "7 = (exchange)" _
& vbNewLine & "8 = (lease)" _
& vbNewLine & "9 = (unlease)" _
& vbNewLine & "10 = (alias)" _
& vbNewLine & "11 = (mass)" _
& vbNewLine & "12 = (data)" _
& vbNewLine & "13 = (smart_account)" _
& vbNewLine & "14 = (sponsorship)" _
& vbNewLine & "141 = (sponsor)" _
& vbNewLine & "15 = (smart_asset)" _
& vbNewLine & "16 = (invoke)" _
& vbNewLine & "161 = (invoke_transfer)" _
& vbNewLine & "162 = (invoke_data)" _
, "tx filter", "7")

Select Case txtype
    Case 0
        txtype = "(fees)"
    Case 1
        txtype = "(genesis)"
    Case 2
        txtype = "(payment)"
    Case 3
        txtype = "(issue)"
    Case 4
        txtype = "(transfer)"
    Case 5
        txtype = "(reissue)"
    Case 6
        txtype = "(burn)"
    Case 7
        txtype = "(exchange)"
    Case 8
        txtype = "(lease)"
    Case 9
        txtype = "(unlease)"
    Case 10
        txtype = "(alias)"
    Case 11
        txtype = "(mass)"
    Case 12
        txtype = "(data)"
    Case 13
        txtype = "(smart_account)"
    Case 14
        txtype = "(sponsorship)"
    Case 141
        txtype = "(sponsor)"
    Case 15
        txtype = "(smart_asset)"
    Case 16
        txtype = "(invoke)"
    Case 161
        txtype = "(invoke_transfer)"
    Case 162
        txtype = "(invoke_data)"
    Case Else
        txtype = "(" & txtype & ")"
End Select

n = Cells(Rows.Count, 4).End(xlUp).Row
For i = 1 To n
Repeat:
    If Cells(i, 4).Value <> "" And Cells(i, 4).Value <> txtype Then
        Rows(i & ":" & i).Select
        Selection.Delete Shift:=xlUp
        GoTo Repeat
    End If
    If Cells(i, 5).Value <> mainA And Cells(i, 7).Value = mainA Then
        s = Cells(i, 5).Value
        Cells(i, 5).Value = Cells(i, 7).Value
        Cells(i, 7).Value = s
        Cells(i, 6).Value = "<-"
    End If
Next
    
'sorting by addresses
Columns("A:H").Select
Selection.Sort Key1:=Range("G1"), Order1:=xlAscending, Key2:=Range("H1") _
, Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False _
, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
:=xlSortTextAsNumbers
        
'check the 1st raw
n = Cells(Rows.Count, 7).End(xlUp).Row
summ = Cells(1, 8).Value
tsumm = summ
na = 1
k = 1
If Cells(2, 7).Value <> Cells(1, 7).Value Then
    Cells(k, 9).Value = na
    Cells(k, 10).Value = Cells(1, 7).Value
    Cells(k, 11).Value = Fix(summ)
    Cells(k, 12).Value = summ - Fix(summ)
End If
GoSub randomColor
Cells(1, 7).Interior.Color = RGB(r, g, B)

'check the other raws
For i = 2 To n
    prevA = Cells(i - 1, 7).Value
    currA = Cells(i, 7).Value
    nextA = Cells(i + 1, 7).Value
    If currA = prevA Then
        summ = summ + Cells(i, 8).Value
        na = na + 1
        Cells(i, 7).Interior.Color = RGB(r, g, B)
    Else
        summ = Cells(i, 8).Value
        na = 1
        GoSub randomColor
        Cells(i, 7).Interior.Color = RGB(r, g, B)
    End If
    If nextA <> currA Then
        k = k + 1
        Cells(k, 9).Value = na
        'if not address and not alias
        If Left(currA, 2) <> "3P" And Asc(Left(currA, 1)) < 91 Then
            Cells(k, 10).Value = currA & " " & summ
            tsumm = tsumm - summ
        Else
            Cells(k, 10).Value = currA
            Cells(k, 11).Value = Fix(summ)
            Cells(k, 12).Value = summ - Fix(summ)
        End If
    End If
    tsumm = tsumm + Cells(i, 8).Value
Next

'sorting by size
Columns("I:L").Select
Selection.Sort Key1:=Range("K1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1 _
, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers

Cells(k + 2, 10).Value = "total addresses"
Cells(k + 2, 11).Value = "total summ"
Cells(k + 3, 10).Value = k
Cells(k + 3, 11).Value = Fix(tsumm)
Cells(k + 3, 12).Value = tsumm - Fix(tsumm)
Cells(k + 5, 10).Value = "total tx " & txtype
Cells(k + 6, 10).Value = n
Columns("A:Z").Select
Selection.Columns.AutoFit
Range("K" & k + 2 & ":" & "L" & k + 2).Select
Selection.Merge

Exit Sub

randomColor:
    r = WorksheetFunction.RandBetween(150, 255)
    g = WorksheetFunction.RandBetween(150, 255)
    B = WorksheetFunction.RandBetween(150, 255)
Return

Application.UseSystemSeparators = True

End Sub
