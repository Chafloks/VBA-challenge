Sub Dosomething()
    Dim xSh As Worksheet
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
    Next
End Sub
	
Sub stocks()

Dim lrow As Integer
Dim count As Integer
Dim found As Integer
Dim op As Double
Dim cl As Double
Dim x As Integer
Dim a As Double

lrow1 = Cells(Rows.count, 1).End(xlUp).row

Cells(1, 11).Value = "Ticker"
Cells(1, 12).Value = "Yearly change"
Cells(1, 13).Value = "% of yearly change"
Cells(1, 14).Value = "Total Volume"
Cells(1, 16).Value = "Greatest % Increase"
Cells(1, 17).Value = "Greatest % Deccrease"
Cells(1, 18).Value = "Greatest Total Volume"


For i = 2 To lrow1
        found = 0
        For j = 2 To count + 2
            If Cells(i, 1).Value = Cells(j, 11).Value Then
                found = 1
                Cells(j, 14).Value = Cells(j, 14).Value + Cells(i, 7).Value
                x = 1
                Exit For
            End If
        Next j
        If found = 0 Then
            If x = 1 Then
                cl = Cells(i - 1, 6)
                Cells(count + 1, 12).Value = cl - op
				
                If Cells(count + 1, 12).Value > 0 Then
                    Cells(count + 1, 12).Interior.ColorIndex = 4
                Else
                    Cells(count + 1, 12).Interior.ColorIndex = 3
                End If
                If op = 0 Then ' if denominator equals 0 then division by 0 occurs
                    Cells(count + 1, 13).Value = Null
                Else
                    Cells(count + 1, 13).Value = ((Cells(count + 1, 12).Value / op) * 100) & "%"
                End If
            End If
            Cells(count + 2, 11) = Cells(i, 1).Value
            Cells(count + 2, 14) = Cells(i, 7).Value
            count = count + 1
            op = Cells(i, 3)
        End If
Next i
g
lrow2 = Cells(Rows.count, 11).End(xlUp).row
cl = Cells(lrow1, 6).Value
Cells(lrow2, 12).Value = cl - op
If Cells(lrow2, 12).Value > 0 Then
    Cells(lrow2, 12).Interior.ColorIndex = 4
Else
    Cells(lrow2, 12).Interior.ColorIndex = 3
End If
Cells(lrow2, 13).Value = ((Cells(lrow2, 12).Value / op) * 100) & "%"
a = Application.WorksheetFunction.Max(Range("M:M"))
Cells(2, 16).Value = a * 100 & "%"
a = Application.WorksheetFunction.Min(Range("M:M"))
Cells(2, 17).Value = a * 100 & "%"
a = Application.WorksheetFunction.Max(Range("N:N"))
Cells(2, 18).Value = a
End Sub
