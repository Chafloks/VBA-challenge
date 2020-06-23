Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub

Sub Dosomething()
    Dim xSh As Worksheet
    'Optimize Code
	Call OptimizeCode_Begin
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
    Next
    'Optimize Code
	Call OptimizeCode_End
End Sub
Sub stocks()

Dim lrow As Long
Dim count As Long
Dim row As Integer
Dim found As Integer
Dim op As Double
Dim cl As Double
Dim x As Long
Dim a As Double

lrow1 = Cells(Rows.count, 1).End(xlUp).row
row = 1

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
            ElseIf Cells(i + 1, 2) = Null Then
                found = 0
            End If
        Next j
        If Cells(i, 2) = Null Then
            found = 0
        End If
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
