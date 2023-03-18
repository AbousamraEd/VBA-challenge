Attribute VB_Name = "Module2"
Sub Functionality()
Cells(2, 12).Value = "Greatest % Increase"
Cells(3, 12).Value = "Greatest % Decrease"
Cells(4, 12).Value = "Greatest Total Volume"
Cells(1, 13).Value = "Ticker"
Cells(1, 14).Value = "Value"

Dim rowcount As Long
rowcount = Cells(Rows.Count, 1).End(xlUp).Row


Cells(2, 14).Value = WorksheetFunction.Max(Range(Cells(2, 10), Cells(rowcount, 10)))
Cells(3, 14).Value = WorksheetFunction.Min(Range(Cells(2, 10), Cells(rowcount, 10)))
Cells(4, 14).Value = WorksheetFunction.Max(Range(Cells(2, 11), Cells(rowcount, 11)))

For i = 2 To rowcount
    If Cells(i, 10).Value = WorksheetFunction.Max(Range(Cells(2, 10), Cells(rowcount, 10))) Then
    Cells(2, 13).Value = Cells(i, 8).Value
    End If
    rowcount = rowcount + 1
Next i

For i = 2 To rowcount
    If Cells(i, 10).Value = WorksheetFunction.Min(Range(Cells(2, 10), Cells(rowcount, 10))) Then
    Cells(3, 13).Value = Cells(i, 8).Value
    End If
    rowcount = rowcount + 1
Next i

For i = 2 To rowcount
    If Cells(i, 11).Value = WorksheetFunction.Max(Range(Cells(2, 11), Cells(rowcount, 11))) Then
    Cells(4, 13).Value = Cells(i, 8).Value
    End If
    rowcount = rowcount + 1
Next i

Sub Functionality()
Cells(2, 12).Value = "Greatest % Increase"
Cells(3, 12).Value = "Greatest % Decrease"
Cells(4, 12).Value = "Greatest Total Volume"
Cells(1, 13).Value = "Ticker"
Cells(1, 14).Value = "Value"

Dim rowcount As Long
rowcount = Cells(Rows.Count, 1).End(xlUp).Row


Cells(2, 14).Value = WorksheetFunction.Max(Range(Cells(2, 10), Cells(rowcount, 10)))
Cells(3, 14).Value = WorksheetFunction.Min(Range(Cells(2, 10), Cells(rowcount, 10)))
Cells(4, 14).Value = WorksheetFunction.Max(Range(Cells(2, 11), Cells(rowcount, 11)))

For i = 2 To rowcount
    If Cells(i, 10).Value = WorksheetFunction.Max(Range(Cells(2, 10), Cells(rowcount, 10))) Then
    Cells(2, 13).Value = Cells(i, 8).Value
    End If
    rowcount = rowcount + 1
Next i

For i = 2 To rowcount
    If Cells(i, 10).Value = WorksheetFunction.Min(Range(Cells(2, 10), Cells(rowcount, 10))) Then
    Cells(3, 13).Value = Cells(i, 8).Value
    End If
    rowcount = rowcount + 1
Next i

For i = 2 To rowcount
    If Cells(i, 11).Value = WorksheetFunction.Max(Range(Cells(2, 11), Cells(rowcount, 11))) Then
    Cells(4, 13).Value = Cells(i, 8).Value
    End If
    rowcount = rowcount + 1
Next i


End Sub


End Sub
