Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
    Next
    Application.ScreenUpdating = True
End Sub
Sub stocks()

Dim lrow1 As Long
Dim lrow2 As Long
Dim count As Integer
Dim found As Integer
Dim tk As String
Dim tv As Double
Dim op As Double
Dim cl As Double
Dim diff As Double
Dim perc As Double
Dim x As Long
Dim y As Integer

y = 2
lrow1 = Cells(Rows.count, 1).End(xlUp).row

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly change"
Cells(1, 11).Value = "% of yearly change"
Cells(1, 12).Value = "Total Volume"
Cells(1, 13).Value = "Greatest % Increase"
Cells(1, 14).Value = "Greatest % Deccrease"
Cells(1, 15).Value = "Greatest Total Volume"



For i = 2 To lrow1
    If Cells(i, 1) <> Cells(i + 1, 1) Then
        tk = Cells(i, 1).Value
        tv = Cells(i, 7).Value + tv
        op = Cells(i - x, 3).Value
        cl = Cells(i, 6).Value
        diff = cl - op
        If op = 0 Then
            perc = Null
        Else
            perc = diff / op
        End If
        Range("I" & y).Value = tk
        Range("J" & y).Value = diff
        Range("K" & y).Value = perc
        Range("L" & y).Value = tv
        y = y + 1
        
        tv = 0
        tk = 0
        op = 0
        diff = 0
        perc = 0
        x = 0
    Else
        tv = Cells(i, 7).Value + tv
        x = x + 1
    End If
Next i
lrow2 = Cells(Rows.count, 9).End(xlUp).row
For i = 2 To lrow2
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    Else
        Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
a = Application.WorksheetFunction.Max(Range("K:K"))
Cells(2, 13).Value = a * 100 & "%"
a = Application.WorksheetFunction.Min(Range("K:K"))
Cells(2, 14).Value = a * 100 & "%"
a = Application.WorksheetFunction.Max(Range("L:L"))
Cells(2, 15).Value = a

End Sub

