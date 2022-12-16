Attribute VB_Name = "Module1"
Sub stockex()
'Set variables
Dim days20, i, imax, di, j As Integer
Dim lastrow As Long
'Determine number of days in trading year
If Cells(2, 2).Value = 20200102 Then
    days20 = 253
    ElseIf Cells(2, 2).Value = 20180102 Then
    days20 = 251
    Else: days20 = 252
    End If
'Determine number of stocks in sheet
lastrow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
imax = ((lastrow - 1) / days20) - 1
For i = 0 To imax
    di = (days20) * i
'Pull ticker symbol and calculate changes in stocks.
Cells(2 + i, 9).Value = Cells(2 + di, 1).Value
Cells(2 + i, 10).Value = (Cells(254 + di, 6).Value - Cells(2 + di, 3).Value)
If Cells(2 + di, 3).Value = 0 Then
    Cells(2 + i, 11).Value = 0
    Else: Cells(2 + i, 11).Value = Cells(2 + i, 10).Value / Cells(2 + di, 3).Value
    End If
'Initialize sum in case script is ran twice.
Cells(2 + i, 12).Value = 0
For j = 0 To (days20 - 1)
Cells(2 + i, 12).Value = Cells(2 + i, 12).Value + Cells(2 + j + di, 7)
Next j
Next i
'Add conditional formatting to necessary cells
Dim colors As Range
Dim positive As FormatCondition, negative As FormatCondition
Set colors = Range("J2:J3500")
Set positive = colors.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set negative = colors.FormatConditions.Add(xlCellValue, xlLess, "=0")
With positive
    .Interior.Color = vbGreen
   End With
With negative
    .Interior.Color = vbRed
   End With
'find maximum values
Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K:K"))
Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K:K"))
Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L:L"))
'Print tickers of maximum values
Dim range1, range2, range3, dummyrange As Range
range1 = Range("K:K").Find(Range("Q2")).Address
range2 = Range("K:K").Find(Range("Q3")).Address
range3 = Range("L:L").Find(Range("Q4")).Address
Cells(2, 16).Value = range1
Cells(3, 16).Value = range2
Cells(4, 16).Value = range3
'Change number format to percentages
Range("K:K").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"

'Label newly created columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Location of Value"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Trade Volume"
'Auto fit columns to make it look nicer
Columns("A:Q").AutoFit
End Sub
