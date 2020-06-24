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
dim tk as string
dim tv as long
Dim op As Double
Dim cl As Double
dim diff as double
dim perc as double
dim x as integer
dim y as integer


lrow1 = Cells(Rows.count, 1).End(xlUp).row

Cells(1, 11).Value = "Ticker"
Cells(1, 12).Value = "Yearly change"
Cells(1, 13).Value = "% of yearly change"
Cells(1, 14).Value = "Total Volume"
Cells(1, 16).Value = "Greatest % Increase"
Cells(1, 17).Value = "Greatest % Deccrease"
Cells(1, 18).Value = "Greatest Total Volume"



For i = 2 To lrow1
	if cells(i, 1)<>cells(i+1,1) then
		tk = ceslls(i,1).value
		tv = cells(i,7).value + tv
		op = cells(i-x,3).value
		cl = cells(i,6).value
		diff = cl - op
		perc = diff / op
		Range("I"& y).value = tk
		Range("J"& y).value = diff
		Range("K"& y).value = perc * 100 & "%"
		Range("L"& y).value = tv
		y = y + 1
		
		tv = 0
		tk = 0
		op = 0
		diff = 0
		perc = 0
		x = 0
	else
		tv = tv + cells(i,7).value
		x= x +1
	end if
Next i

End Sub
