Attribute VB_Name = "Module1"
Sub Ticker()
For Each ws In Worksheets
ws.Cells(1, 9).Value = "Ticker"
Dim rowcount As Long
Dim tickercount As Long
rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
tickercount = 2
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
Dim Var1 As Double
Dim Var2 As Double
Dim Var3 As Double
'var1 is open vars 2 is close
Var1 = ws.Cells(2, 3).Value
Var2 = 0
Var3 = 0

For i = 2 To rowcount
    Var3 = Var3 + ws.Cells(i, 7).Value
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
    ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
    Var2 = ws.Cells(i, 6).Value
    ws.Cells(tickercount, 10).Value = Var2 - Var1
        If ws.Cells(tickercount, 10).Value > 0 Then
        ws.Cells(tickercount, 10).Interior.Color = vbGreen
        Else
        ws.Cells(tickercount, 10).Interior.Color = vbRed
        End If
    ws.Cells(tickercount, 11).Value = FormatPercent((Var2 / Var1) - 1)
    ws.Cells(tickercount, 12).Value = Var3
    Var1 = ws.Cells(i + 1, 3).Value
    Var3 = 0
    tickercount = tickercount + 1
    End If
Next i
Next ws
End Sub




