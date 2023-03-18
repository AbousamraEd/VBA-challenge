Attribute VB_Name = "Module1"
Sub Ticker()
Cells(1, 8).Value = "Ticker"
Dim rowcount As Long
Dim tickercount As Long
rowcount = Cells(Rows.Count, 1).End(xlUp).Row
tickercount = 2
For i = 2 To rowcount
    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    Cells(tickercount, 8).Value = Cells(i, 1).Value
    tickercount = tickercount + 1
    End If
    rowcount = rowcount + 1
Next i
End Sub
Sub YearlyPrice()
Cells(1, 9).Value = "Yearly Change"
Cells(1, 10).Value = "Percent Change"
Cells(1, 11).Value = "Total Stock Volume"
Dim rowcount As Long
Dim tickercount As Long
rowcount = Cells(Rows.Count, 1).End(xlUp).Row
tickercount = 2
For i = 2 To rowcount
    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    Cells(tickercount, 9).Value = Cells(i + 1, 6).Value - Cells(i - 252, 3).Value
    Cells(tickercount, 10).Value = (Cells(i, 6).Value / Cells(i - 252, 3).Value) - 1
    Cells(tickercount, 11).Value = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(i - 252, 7)))
    tickercount = tickercount + 1
    End If
    rowcount = rowcount + 1
Next i
End Sub



