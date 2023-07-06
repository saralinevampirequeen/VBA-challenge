Attribute VB_Name = "Module2"
Sub StockMarketAnalysis2019()
    ' Declares the variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Sets the worksheet to analyze
    Set ws = ThisWorkbook.Sheets("2019")
    
    ' Finds the last row of data in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Sets initial values for variables
    summaryRow = 2
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
        
    ' Adds the headers for the summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Loops through the stock market data
    For i = 2 To lastRow
        ' Checks if it's a new ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' Sets the ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            ' Sets the opening price
            openingPrice = ws.Cells(i, 3).Value
        End If
        
        ' Sets the closing price
        closingPrice = ws.Cells(i, 6).Value
        
        ' Calculates the yearly change and the percent change
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentChange = yearlyChange / openingPrice
        Else
            percentChange = 0
        End If
        
        ' Adds the yearly change, percent change, and total volume to the summary table
        ws.Cells(summaryRow, 9).Value = ticker
        ws.Cells(summaryRow, 10).Value = yearlyChange
        ws.Cells(summaryRow, 11).Value = percentChange
        ws.Cells(summaryRow, 12).Value = ws.Cells(i, 7).Value
        
        ' Adds the conditional formatting for the positive and negative yearly changes
        If yearlyChange >= 0 Then
            ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
        Else
            ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End If
        
        ' Updates the total stock volume for the ticker
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        ' Checks if it's the last occurrence of the ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ' Adds the total volume for the ticker to the summary table
            ws.Cells(summaryRow, 12).Value = totalVolume
            
            ' Resets the total volume for the next ticker
            totalVolume = 0
            
            ' Increments the summary row
            summaryRow = summaryRow + 1
        End If
        
     ' Adds the headers for the greatest values
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
        
        ' Checks for the stock with the greatest percentage increase
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            maxIncreaseTicker = ticker
        End If
        
        ' Checks for the stock with the greatest percentage decrease
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxDecreaseTicker = ticker
        End If
        
        ' Checks for the stock with the greatest total volume
        If ws.Cells(i, 7).Value > maxVolume Then
            maxVolume = ws.Cells(i, 7).Value
            maxVolumeTicker = ticker
        End If
    Next i
    
    ' Adds the greatest values to the summary table
    ws.Cells(2, 16).Value = maxIncreaseTicker
    ws.Cells(2, 17).Value = maxIncrease
    ws.Cells(3, 16).Value = maxDecreaseTicker
    ws.Cells(3, 17).Value = maxDecrease
    ws.Cells(4, 16).Value = maxVolumeTicker
    ws.Cells(4, 17).Value = maxVolume
    
    ' Formats the percentage change column as percentages
    ws.Columns("K").NumberFormat = "0.00%"
    
    ' Autofit columns in the worksheet
    ws.Columns.AutoFit
End Sub


