Attribute VB_Name = "Module1"

Sub CalculateStockDataForAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Long
    Dim openColumn As Long
    Dim closeColumn As Long
    Dim volumeColumn As Long
    Dim currentRow As Long
    Dim currentTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    
    ' Initialize variables
    maxPercentIncrease = -99999999
    maxPercentDecrease = 99999999
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet name is numeric
        If IsNumeric(ws.Name) Then
            ' Define the columns where data is stored
    
            tickerColumn = 1 ' Column A for ticker
            openColumn = 3 ' Column C for open price
            closeColumn = 6 ' Column F for close price
            volumeColumn = 7 ' Column G for volume
            
            ' Find the last row of data
            lastRow = ws.Cells(ws.Rows.Count, tickerColumn).End(xlUp).Row
            
            ' Initialize summary row
            summaryRow = 2
            
            ' Reset variables for each sheet
            maxPercentIncrease = -99999999
            maxPercentDecrease = 99999999
            
            ' Loop through each row of data
            For currentRow = 2 To lastRow ' Assuming your data starts from row 2
                ' Get the current ticker
                currentTicker = ws.Cells(currentRow, tickerColumn).Value
                
                ' If it's a new ticker, calculate and output summary data
                If ws.Cells(currentRow, tickerColumn).Value <> ws.Cells(currentRow - 1, tickerColumn).Value Or currentRow = 2 Then
                    ' Output summary data
                    ws.Cells(summaryRow, 9).Value = currentTicker
                    
                    ' Initialize open price for the ticker
                    openPrice = ws.Cells(currentRow, openColumn).Value
                    
                    ' Initialize total volume
                    totalVolume = 0
                End If
                
                ' Sum up the total volume for the current ticker
                totalVolume = totalVolume + ws.Cells(currentRow, volumeColumn).Value
                
                ' Get the close price for the ticker (use the last instance)
                closePrice = ws.Cells(currentRow, closeColumn).Value
                
                ' If it's the last row for the ticker, calculate yearly and percent change
                If currentRow = lastRow Or ws.Cells(currentRow, tickerColumn).Value <> ws.Cells(currentRow + 1, tickerColumn).Value Then
                    ' Calculate yearly change
                    yearlyChange = closePrice - openPrice
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    
                    ' Calculate percent change
                    If openPrice <> 0 Then
                        percentChange = yearlyChange / openPrice
                    Else
                        percentChange = 0
                    End If
                    ws.Cells(summaryRow, 11).Value = percentChange
                    
                    ' Output total volume for the current ticker
                    ws.Cells(summaryRow, 12).Value = totalVolume
                    
                    ' Check for greatest percent increase, decrease, and total volume
                    If percentChange > maxPercentIncrease Then
                        maxPercentIncrease = percentChange
                        maxPercentIncreaseTicker = currentTicker
                    End If
                    
                    If percentChange < maxPercentDecrease Then
                        maxPercentDecrease = percentChange
                        maxPercentDecreaseTicker = currentTicker
                    End If
                    
                    If totalVolume > maxTotalVolume Then
                        maxTotalVolume = totalVolume
                        maxTotalVolumeTicker = currentTicker
                    End If
                    
                    ' Highlight yearly change cells
                    If yearlyChange > 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                    ElseIf yearlyChange < 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                    Else
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = xlNone ' No color
                    End If
                    
                    ' Move to the next summary row
                    summaryRow = summaryRow + 1
                End If
            Next currentRow
            
            ' Output tickers with greatest percent increase, decrease, and total volume
            With ws
                .Range("O2").Value = "Greatest % Increase"
                .Range("O3").Value = "Greatest % Decrease"
                .Range("O4").Value = "Greatest Total Volume"
                
                .Range("P2").Value = maxPercentIncreaseTicker
                .Range("P3").Value = maxPercentDecreaseTicker
                .Range("P4").Value = maxTotalVolumeTicker
                
                .Range("Q2").Value = maxPercentIncrease
                .Range("Q3").Value = maxPercentDecrease
                .Range("Q4").Value = maxTotalVolume
            End With
        End If
    Next ws
End Sub

