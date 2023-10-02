VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()

'Create new worksheet for table
Sheets.Add.Name = "Summary"

'Name columns in Summary Sheet

Worksheets("Summary").Range("A1").Value = "Ticker"
Worksheets("Summary").Range("B1").Value = "Yearly Change"
Worksheets("Summary").Range("B1").EntireColumn.ColumnWidth = 12
Worksheets("Summary").Range("C1").Value = "Percent Change"
Worksheets("Summary").Range("C1").EntireColumn.ColumnWidth = 13
Worksheets("Summary").Range("D1").Value = "Total Stock Volume"
Worksheets("Summary").Range("D1").EntireColumn.ColumnWidth = 19

Worksheets("Summary").Range("J1").Value = "Ticker"
Worksheets("Summary").Range("K1").Value = "Yearly Change"
Worksheets("Summary").Range("K1").EntireColumn.ColumnWidth = 12
Worksheets("Summary").Range("L1").Value = "Percent Change"
Worksheets("Summary").Range("L1").EntireColumn.ColumnWidth = 13
Worksheets("Summary").Range("M1").Value = "Total Stock Volume"
Worksheets("Summary").Range("M1").EntireColumn.ColumnWidth = 19

Worksheets("Summary").Range("S1").Value = "Ticker"
Worksheets("Summary").Range("T1").Value = "Yearly Change"
Worksheets("Summary").Range("T1").EntireColumn.ColumnWidth = 12
Worksheets("Summary").Range("U1").Value = "Percent Change"
Worksheets("Summary").Range("U1").EntireColumn.ColumnWidth = 13
Worksheets("Summary").Range("V1").Value = "Total Stock Volume"
Worksheets("Summary").Range("V1").EntireColumn.ColumnWidth = 19


'___________________________________________________________________________________________________

'Loop through each worksheet for values

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

Dim ticker As String
Dim pricedate As String
Dim openprice As Double
Dim closeprice As Double
Dim stockvolume As LongLong
Dim i As Long
Dim j As Long
Dim k As Long
Dim totalstockvol As LongLong
totalstockvol = 0
ticker = " "

If ws.Name = "2018" Then

'Determine last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row of worksheet and pull/store data
For i = 2 To LastRow

    ticker = ws.Cells(i, 1).Value
    pricedate = ws.Cells(i, 2).Value
    openprice = ws.Cells(i, 3).Value
    closeprice = ws.Cells(i, 6).Value
    stockvolume = ws.Cells(i, 7).Value
    
Dim yrlyopenprice As Double
Dim yrlycloseprice As Double

    If Cells(i + 1, 1).Value = Cells(i, 1).Value And (pricedate = "20180102") Then
    'Set this as the yearly opening price
    yrlyopenprice = openprice
    'Print ticker in Summary Sheet
    Worksheets("Summary").Cells(Rows.Count, 1).End(xlUp).Offset(1) = ticker
    
    ElseIf Cells(i - 1, 1).Value = Cells(i, 1).Value And (pricedate = "20181231") Then
    'Set this as the yearly closing price
    yrlycloseprice = closeprice
    Dim yrlychange As Double
    
    'yearly change = closing price - opening price
    Worksheets("Summary").Cells(Rows.Count, 2).End(xlUp).Offset(1) = yrlycloseprice - yrlyopenprice
    
    'percentage of opening price - yearly change/opening price *100
    Worksheets("Summary").Cells(Rows.Count, 3).End(xlUp).Offset(1) = FormatPercent(((yrlycloseprice - yrlyopenprice) / yrlyopenprice), 2)
    End If
    
    'total stock volume calculation
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    totalstockvol = totalstockvol + stockvolume
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    totalstockvol = totalstockvol + stockvolume
    
    'Print totalstockvol in other sheet
    Worksheets("Summary").Cells(Rows.Count, 4).End(xlUp).Offset(1) = totalstockvol
    totalstockvol = 0
    End If
    
 Next i

ElseIf ws.Name = "2019" Then
'Determine last row
LastRow19 = Worksheets("2019").Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row of worksheet and pull/store data
For j = 2 To LastRow19

    ticker = Worksheets("2019").Cells(j, 1).Value
    pricedate = Worksheets("2019").Cells(j, 2).Value
    openprice = Worksheets("2019").Cells(j, 3).Value
    closeprice = Worksheets("2019").Cells(j, 6).Value
    stockvolume = Worksheets("2019").Cells(j, 7).Value

    If Worksheets("2019").Cells(j + 1, 1).Value = Worksheets("2019").Cells(j, 1).Value And (pricedate = "20190102") Then
    'Set this as the yearly opening price
    yrlyopenprice = openprice
    'Print ticker in Summary Sheet
    Worksheets("Summary").Cells(Rows.Count, 10).End(xlUp).Offset(1) = ticker
    
    ElseIf Worksheets("2019").Cells(j - 1, 1).Value = Worksheets("2019").Cells(j, 1).Value And (pricedate = "20191231") Then
    'Set this as the yearly closing price
    yrlycloseprice = closeprice
    'Dim yrlychange As Double
    
    'yearly change = closing price - opening price
    Worksheets("Summary").Cells(Rows.Count, 11).End(xlUp).Offset(1) = yrlycloseprice - yrlyopenprice
    
    'percentage of opening price - yearly change/opening price *100
    Worksheets("Summary").Cells(Rows.Count, 12).End(xlUp).Offset(1) = FormatPercent(((yrlycloseprice - yrlyopenprice) / yrlyopenprice), 2)
    End If
    
    'total stock volume calculation
    If Cells(j + 1, 1).Value = Cells(j, 1).Value Then
    totalstockvol = totalstockvol + stockvolume
    ElseIf Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
    totalstockvol = totalstockvol + stockvolume
    
    'Print totalstockvol in other sheet
    Worksheets("Summary").Cells(Rows.Count, 13).End(xlUp).Offset(1) = totalstockvol
    totalstockvol = 0
    End If

Next j


ElseIf ws.Name = "2020" Then
'Determine last row
LastRow20 = Worksheets("2020").Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row of worksheet and pull/store data
For k = 2 To LastRow20

    ticker = Worksheets("2020").Cells(k, 1).Value
    pricedate = Worksheets("2020").Cells(k, 2).Value
    openprice = Worksheets("2020").Cells(k, 3).Value
    closeprice = Worksheets("2020").Cells(k, 6).Value
    stockvolume = Worksheets("2020").Cells(k, 7).Value

    If Worksheets("2020").Cells(k + 1, 1).Value = Worksheets("2020").Cells(k, 1).Value And (pricedate = "20200102") Then
    'Set this as the yearly opening price
    yrlyopenprice = openprice
    'Print ticker in Summary Sheet
    Worksheets("Summary").Cells(Rows.Count, 19).End(xlUp).Offset(1) = ticker
    
    ElseIf Worksheets("2020").Cells(k - 1, 1).Value = Worksheets("2020").Cells(k, 1).Value And (pricedate = "20201231") Then
    'Set this as the yearly closing price
    yrlycloseprice = closeprice
    'Dim yrlychange As Double
    
    'yearly change = closing price - opening price
    Worksheets("Summary").Cells(Rows.Count, 20).End(xlUp).Offset(1) = yrlycloseprice - yrlyopenprice
    
    'percentage of opening price - yearly change/opening price *100
    Worksheets("Summary").Cells(Rows.Count, 21).End(xlUp).Offset(1) = FormatPercent(((yrlycloseprice - yrlyopenprice) / yrlyopenprice), 2)
    End If
    
    'total stock volume calculation
    If Cells(k + 1, 1).Value = Cells(k, 1).Value Then
    totalstockvol = stockvolume + totalstockvol
    ElseIf Cells(k + 1, 1).Value <> Cells(k, 1).Value Then
    stockvolume = Worksheets("2020").Cells(k, 7).Value
    totalstockvol = totalstockvol + stockvolume
    
    'Print totalstockvol in other sheet
    Worksheets("Summary").Cells(Rows.Count, 22).End(xlUp).Offset(1) = totalstockvol
    totalstockvol = 0
    End If
    
Next k

End If

Next ws

'_________________________________________________________________________________________

'Conditional formatting yearly change cells red for <0, green for >=0

SumLastRow = Worksheets("Summary").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To SumLastRow
    If Worksheets("Summary").Cells(i, 2).Value >= 0 Then
    Worksheets("Summary").Cells(i, 2).Interior.ColorIndex = 4
    
    ElseIf Worksheets("Summary").Cells(i, 2).Value < 0 Then
    Worksheets("Summary").Cells(i, 2).Interior.ColorIndex = 3
    
    End If
    
    If Worksheets("Summary").Cells(i, 3).Value >= 0 Then
    Worksheets("Summary").Cells(i, 3).Interior.ColorIndex = 4
    
    ElseIf Worksheets("Summary").Cells(i, 3).Value < 0 Then
    Worksheets("Summary").Cells(i, 3).Interior.ColorIndex = 3
    
    End If
    
    '------
    If Worksheets("Summary").Cells(i, 11).Value >= 0 Then
    Worksheets("Summary").Cells(i, 11).Interior.ColorIndex = 4
    
    ElseIf Worksheets("Summary").Cells(i, 11).Value < 0 Then
    Worksheets("Summary").Cells(i, 11).Interior.ColorIndex = 3
    
    End If
    
    If Worksheets("Summary").Cells(i, 12).Value >= 0 Then
    Worksheets("Summary").Cells(i, 12).Interior.ColorIndex = 4
    
    ElseIf Worksheets("Summary").Cells(i, 12).Value < 0 Then
    Worksheets("Summary").Cells(i, 12).Interior.ColorIndex = 3
    
    End If
    
    '------
    If Worksheets("Summary").Cells(i, 20).Value >= 0 Then
    Worksheets("Summary").Cells(i, 20).Interior.ColorIndex = 4
    
    ElseIf Worksheets("Summary").Cells(i, 20).Value < 0 Then
    Worksheets("Summary").Cells(i, 20).Interior.ColorIndex = 3
    
    End If
    
    If Worksheets("Summary").Cells(i, 21).Value >= 0 Then
    Worksheets("Summary").Cells(i, 21).Interior.ColorIndex = 4
    
    ElseIf Worksheets("Summary").Cells(i, 21).Value < 0 Then
    Worksheets("Summary").Cells(i, 21).Interior.ColorIndex = 3
    
    End If
    
Next i


'label for rows, column for ticker, column for value

Worksheets("Summary").Range("F2").Value = "Greatest % Increase"

Worksheets("Summary").Range("F3").Value = "Greatest % Decrease"

Worksheets("Summary").Range("F4").Value = "Greatest Total Volume"

Worksheets("Summary").Range("F1").EntireColumn.ColumnWidth = 20

Worksheets("Summary").Range("G1").Value = "Ticker"
Worksheets("Summary").Range("H1").Value = "Value"
Worksheets("Summary").Range("H1").EntireColumn.ColumnWidth = 15

Worksheets("Summary").Range("O2").Value = "Greatest % Increase"

Worksheets("Summary").Range("O3").Value = "Greatest % Decrease"

Worksheets("Summary").Range("O4").Value = "Greatest Total Volume"

Worksheets("Summary").Range("O1").EntireColumn.ColumnWidth = 20

Worksheets("Summary").Range("P1").Value = "Ticker"
Worksheets("Summary").Range("Q1").Value = "Value"
Worksheets("Summary").Range("Q1").EntireColumn.ColumnWidth = 15

Worksheets("Summary").Range("X2").Value = "Greatest % Increase"

Worksheets("Summary").Range("X3").Value = "Greatest % Decrease"

Worksheets("Summary").Range("X4").Value = "Greatest Total Volume"

Worksheets("Summary").Range("X1").EntireColumn.ColumnWidth = 20

Worksheets("Summary").Range("Y1").Value = "Ticker"
Worksheets("Summary").Range("Z1").Value = "Value"
Worksheets("Summary").Range("Z1").EntireColumn.ColumnWidth = 15


'Calculate the max/min and ticker values by looping through each row of Summary sheet
Dim max As Double
Dim min As Double
Dim maxincticker As String
Dim maxdecticker As String

Dim maxvol As LongLong
Dim maxvolticker As String

'For 2018
For i = 2 To SumLastRow
    If Worksheets("Summary").Cells(i + 1, 3).Value >= Worksheets("Summary").Cells(i, 3).Value And Worksheets("Summary").Cells(i + 1, 3).Value >= max Then
    max = Worksheets("Summary").Cells(i + 1, 3).Value
    maxincticker = Worksheets("Summary").Cells(i + 1, 1).Value
    ElseIf Worksheets("Summary").Cells(i + 1, 3).Value < Worksheets("Summary").Cells(i, 3).Value And Worksheets("Summary").Cells(i, 3).Value > max Then
    max = Worksheets("Summary").Cells(i, 3).Value
    maxincticker = Worksheets("Summary").Cells(i, 1).Value
    End If
    
    
    If Worksheets("Summary").Cells(i + 1, 3).Value <= Worksheets("Summary").Cells(i, 3).Value And Worksheets("Summary").Cells(i + 1, 3).Value <= min Then
    min = Worksheets("Summary").Cells(i + 1, 3).Value
    maxdecticker = Worksheets("Summary").Cells(i + 1, 1).Value
    ElseIf Worksheets("Summary").Cells(i + 1, 3).Value > Worksheets("Summary").Cells(i, 3).Value And Worksheets("Summary").Cells(i, 3).Value < min Then
    min = Worksheets("Summary").Cells(i, 3).Value
    maxdecticker = Worksheets("Summary").Cells(i, 1).Value
    End If
    
    
    If Worksheets("Summary").Cells(i + 1, 4).Value >= Worksheets("Summary").Cells(i, 4).Value And Worksheets("Summary").Cells(i + 1, 4).Value >= maxvol Then
    maxvol = Worksheets("Summary").Cells(i + 1, 4).Value
    maxvolticker = Worksheets("Summary").Cells(i + 1, 1).Value
    ElseIf Worksheets("Summary").Cells(i + 1, 4).Value < Worksheets("Summary").Cells(i, 4).Value And Worksheets("Summary").Cells(i, 4).Value > maxvol Then
    maxvol = Worksheets("Summary").Cells(i, 4).Value
    maxvolticker = Worksheets("Summary").Cells(i, 1).Value
    
    
    End If
Next i

'Print the max/min and ticker values in assigned cells
Worksheets("Summary").Range("H2").Value = max
Worksheets("Summary").Range("H3").Value = min
Worksheets("Summary").Range("H4").Value = maxvol

Worksheets("Summary").Range("G2").Value = maxincticker
Worksheets("Summary").Range("G3").Value = maxdecticker
Worksheets("Summary").Range("G4").Value = maxvolticker

Worksheets("Summary").Range("H2").NumberFormat = "0.00%"
Worksheets("Summary").Range("H3").NumberFormat = "0.00%"

'For 2019
For j = 2 To SumLastRow
    If Worksheets("Summary").Cells(j + 1, 12).Value >= Worksheets("Summary").Cells(j, 12).Value And Worksheets("Summary").Cells(j + 1, 12).Value >= max Then
    max = Worksheets("Summary").Cells(j + 1, 12).Value
    maxincticker = Worksheets("Summary").Cells(j + 1, 10).Value
    ElseIf Worksheets("Summary").Cells(j + 1, 12).Value < Worksheets("Summary").Cells(j, 12).Value And Worksheets("Summary").Cells(j, 12).Value > max Then
    max = Worksheets("Summary").Cells(j, 12).Value
    maxincticker = Worksheets("Summary").Cells(j, 10).Value
    End If
    
    If Worksheets("Summary").Cells(j + 1, 12).Value <= Worksheets("Summary").Cells(j, 12).Value And Worksheets("Summary").Cells(j + 1, 12).Value <= min Then
    min = Worksheets("Summary").Cells(j + 1, 12).Value
    maxdecticker = Worksheets("Summary").Cells(i + 1, 10).Value
    ElseIf Worksheets("Summary").Cells(j + 1, 12).Value > Worksheets("Summary").Cells(j, 12).Value And Worksheets("Summary").Cells(j, 12).Value < min Then
    min = Worksheets("Summary").Cells(j, 12).Value
    maxdecticker = Worksheets("Summary").Cells(j, 10).Value
    End If
    
    If Worksheets("Summary").Cells(j + 1, 13).Value >= Worksheets("Summary").Cells(j, 13).Value And Worksheets("Summary").Cells(j + 1, 13).Value >= maxvol Then
    maxvol = Worksheets("Summary").Cells(j + 1, 13).Value
    maxvolticker = Worksheets("Summary").Cells(j + 1, 10).Value
    ElseIf Worksheets("Summary").Cells(j + 1, 13).Value < Worksheets("Summary").Cells(j, 13).Value And Worksheets("Summary").Cells(j, 13).Value > maxvol Then
    maxvol = Worksheets("Summary").Cells(j, 13).Value
    maxvolticker = Worksheets("Summary").Cells(j, 10).Value
    
    End If

Next j


'Print the max/min and ticker values in assigned cells
Worksheets("Summary").Range("Q2").Value = max
Worksheets("Summary").Range("Q3").Value = min
Worksheets("Summary").Range("Q4").Value = maxvol

Worksheets("Summary").Range("P2").Value = maxincticker
Worksheets("Summary").Range("P3").Value = maxdecticker
Worksheets("Summary").Range("P4").Value = maxvolticker

Worksheets("Summary").Range("Q2").NumberFormat = "0.00%"
Worksheets("Summary").Range("Q3").NumberFormat = "0.00%"

'For 2020
For k = 2 To SumLastRow
    If Worksheets("Summary").Cells(k + 1, 21).Value >= Worksheets("Summary").Cells(k, 21).Value And Worksheets("Summary").Cells(k + 1, 21).Value >= max Then
    max = Worksheets("Summary").Cells(k + 1, 21).Value
    maxincticker = Worksheets("Summary").Cells(k + 1, 19).Value
    ElseIf Worksheets("Summary").Cells(k + 1, 21).Value < Worksheets("Summary").Cells(k, 21).Value And Worksheets("Summary").Cells(k, 21).Value > max Then
    max = Worksheets("Summary").Cells(k, 21).Value
    maxincticker = Worksheets("Summary").Cells(k, 19).Value
    End If
    
    If Worksheets("Summary").Cells(k + 1, 21).Value <= Worksheets("Summary").Cells(k, 21).Value And Worksheets("Summary").Cells(k + 1, 21).Value <= min Then
    min = Worksheets("Summary").Cells(k + 1, 21).Value
    maxdecticker = Worksheets("Summary").Cells(k + 1, 19).Value
    ElseIf Worksheets("Summary").Cells(k + 1, 21).Value > Worksheets("Summary").Cells(k, 21).Value And Worksheets("Summary").Cells(k, 21).Value < min Then
    min = Worksheets("Summary").Cells(k, 21).Value
    maxdecticker = Worksheets("Summary").Cells(k, 19).Value
    End If
    
    
    If Worksheets("Summary").Cells(k + 1, 22).Value >= Worksheets("Summary").Cells(k, 22).Value And Worksheets("Summary").Cells(k + 1, 22).Value >= maxvol Then
    maxvol = Worksheets("Summary").Cells(k + 1, 22).Value
    maxvolticker = Worksheets("Summary").Cells(k + 1, 19).Value
    ElseIf Worksheets("Summary").Cells(k + 1, 22).Value < Worksheets("Summary").Cells(k, 22).Value And Worksheets("Summary").Cells(k, 22).Value > maxvol Then
    maxvol = Worksheets("Summary").Cells(k, 22).Value
    maxvolticker = Worksheets("Summary").Cells(k, 19).Value
    End If

Next k


'Print the max/min and ticker values in assigned cells
Worksheets("Summary").Range("Z2").Value = max
Worksheets("Summary").Range("Z3").Value = min
Worksheets("Summary").Range("Z4").Value = maxvol

Worksheets("Summary").Range("Y2").Value = maxincticker
Worksheets("Summary").Range("Y3").Value = maxdecticker
Worksheets("Summary").Range("Y4").Value = maxvolticker

Worksheets("Summary").Range("Z2").NumberFormat = "0.00%"
Worksheets("Summary").Range("Z3").NumberFormat = "0.00%"

End Sub

