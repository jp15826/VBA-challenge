Attribute VB_Name = "Module1"
Sub vbaChallenge():

For Each ws In ThisWorkbook.Worksheets
ws.Activate

cellValue = "Ticker"
cellValue1 = "Yearly Change"
cellvalue2 = "Percent Change"
cellvalue3 = "Total Stock Volume"
Range("I1").Value = cellValue
Range("J1").Value = cellValue1
Range("K1").Value = cellvalue2
Range("L1").Value = cellvalue3
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Dim tickerName As String
Dim lastRow As Double
Dim row As Double
Dim totalVolume As Double
    totalVolume = 0
Dim tableRow As Integer
   tableRow = 2
Dim yearlyChange As Double
    yearlyChange = 0
Dim percentChange As Double
Dim firstOpen As Double
    firstOpen = Cells(2, 3).Value
Dim lastClose As Double

    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    For row = 2 To lastRow
       
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
             tickerName = Cells(row, 1).Value
            totalVolume = totalVolume + Cells(row, 7).Value
        
             lastClose = Cells(row, 6).Value
                yearlyChange = lastClose - firstOpen
                percentChange = yearlyChange / firstOpen
            firstOpen = Cells(row + 1, 3).Value
            
                Range("I" & tableRow).Value = Cells(row, 1).Value
                Range("J" & tableRow).Value = yearlyChange
                Range("J" & tableRow).NumberFormat = "0.00"
                Range("K" & tableRow).Value = percentChange
                Range("K" & tableRow).NumberFormat = "0.00%"
                Range("L" & tableRow).Value = totalVolume
          
            totalVolume = 0
            yearlyChange = 0
            tableRow = tableRow + 1
        
        Else
            totalVolume = totalVolume + Cells(row, 7).Value
        End If
    
    Next row
    
    result = WorksheetFunction.Max(Range("K2:K" & tableRow))
    Range("Q2").Value = result
    Range("Q2").NumberFormat = "0.00%"
    result2 = WorksheetFunction.Min(Range("K2:K" & tableRow))
    Range("Q3").Value = result2
    Range("Q3").NumberFormat = "0.00%"
    result = WorksheetFunction.Max(Range("L2:L" & tableRow))
    Range("Q4").Value = result
 
    Range("P2").Value = Application.WorksheetFunction.Index(Range("I1:L" & tableRow), Application.WorksheetFunction.Match(Range("Q2"), Range("K1:K" & tableRow), 0), Application.WorksheetFunction.Match(Range("P1"), Range("I1:L1"), 0))
    Range("P3").Value = Application.WorksheetFunction.Index(Range("I1:L" & tableRow), Application.WorksheetFunction.Match(Range("Q3"), Range("K1:K" & tableRow), 0), Application.WorksheetFunction.Match(Range("P1"), Range("I1:L1"), 0))
    Range("P4").Value = Application.WorksheetFunction.Index(Range("I1:L" & tableRow), Application.WorksheetFunction.Match(Range("Q4"), Range("L1:L" & tableRow), 0), Application.WorksheetFunction.Match(Range("P1"), Range("I1:L1"), 0))
Next ws

End Sub

