Sub StockDataForAllSheets()
    Dim ws As Worksheet
    ' Loop through each worksheet in the workbook
    
    For Each ws In ThisWorkbook.Sheets
        Call StockData(ws)
    Next ws
End Sub
Sub StockData(ws As Worksheet)
    ' Declare variables for processing stock data
    Dim ticker As String
    Dim quarterly_change As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startPrice As Double ' Corrected variable name for consistency
    Dim endPrice As Double
   
    ' Variables to track the greatest increases, decreases, and total volume
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolumeTicker As String
    Dim greatestTotalVolume As Double
   
    ' Initialize variables
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
   
    ' Setup for loop to process each row of stock data
    Dim currentRow As Long
    Dim summaryRow As Long
    Dim lastRow As Long
   
    ' Initialize summary row where first summary begins
    currentRow = 2
    summaryRow = 2
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Correctly find the last row with data
   
    ' Setup headers for the summary output
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Qaurterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
   
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
   
    ' Initialize variables to zero for the loop
    totalVolume = 0
    startPrice = 0
    endPrice = 0
   
    ' Begin looping through each row to process data
    While currentRow <= lastRow
        totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
       
        ' Check if this is the last row for the current ticker
        If ws.Cells(currentRow + 1, 1).Value <> ws.Cells(currentRow, 1).Value Then
            ticker = ws.Cells(currentRow, 1).Value
            endPrice = ws.Cells(currentRow, 6).Value
           
            ' Calculate quarterly change from startPrice to endPrice
            quarterly_change = endPrice - startPrice
            ' Calculate percent change if startPrice is not zero to avoid division by zero
            If startPrice <> 0 Then
                percentChange = (quarterly_change / startPrice) * 100
            Else
                percentChange = 0
            End If
           
            ' Write calculated data to the summary table
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = quarterly_change
            ws.Cells(summaryRow, 11).Value = percentChange & "%"
            ws.Cells(summaryRow, 12).Value = totalVolume
           
            ' Conditional formatting for quarterly_changee
            If quarterly_change < 0 Then
                ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf quarterly_change > 0 Then
                ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                
            End If
           
            ' Update greatest increases, decreases, and volume
            If percentChange > greatestPercentIncrease Then
                greatestPercentIncreaseTicker = ticker
                greatestPercentIncrease = percentChange
            End If
           
            If percentChange < greatestPercentDecrease Then
                greatestPercentDecreaseTicker = ticker
                greatestPercentDecrease = percentChange
            End If
           
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolumeTicker = ticker
                greatestTotalVolume = totalVolume ' Corrected to assign totalVolume
            End If
           
            ' Prepare for next ticker
            summaryRow = summaryRow + 1
            totalVolume = 0
            startPrice = 0
            endPrice = 0
        ElseIf startPrice = 0 Then
            ' Initialize startPrice for the first time we encounter a new ticker
            startPrice = ws.Cells(currentRow, 3).Value
        End If
       
        ' Move to the next row
        currentRow = currentRow + 1
    Wend
   
    ' Write out the greatest increase, decrease, and volume
    ws.Cells(2, 16).Value = greatestPercentIncreaseTicker
    ws.Cells(2, 17).Value = greatestPercentIncrease & "%"
    ws.Cells(3, 16).Value = greatestPercentDecreaseTicker
    ws.Cells(3, 17).Value = greatestPercentDecrease & "%"
    ws.Cells(4, 16).Value = greatestTotalVolumeTicker
    ws.Cells(4, 17).Value = greatestTotalVolume
End Sub
