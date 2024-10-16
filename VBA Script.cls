VERSION 1.0 CLASS

Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTickerSymbol As String
    Dim decreaseTickerSymbol As String
    Dim volumeTickerSymbol As String
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Add headers for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if we're still within the same ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                tickerSymbol = ws.Cells(i, 1).Value
                
                ' Set the close price
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearlyChange = closePrice - openPrice
                
                ' Calculate percent change
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Print the summary
                ws.Cells(summaryRow, 9).Value = tickerSymbol
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format yearly change
                If yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                End If
                
                ' Format percent change
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Check for greatest increase, decrease, and volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerSymbol = tickerSymbol
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerSymbol = tickerSymbol
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerSymbol = tickerSymbol
                End If
                
                ' Reset variables for next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                
                ' Set the opening price for the next ticker
                openPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Set the opening price if it's the first row of the ticker
                If totalVolume = ws.Cells(i, 7).Value Then
                    openPrice = ws.Cells(i, 3).Value
                End If
            End If
        Next i
        
        ' Print the greatest increase, decrease, and volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = increaseTickerSymbol
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = decreaseTickerSymbol
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = volumeTickerSymbol
        ws.Cells(4, 17).Value = greatestVolume
        
    Next ws
End Sub

