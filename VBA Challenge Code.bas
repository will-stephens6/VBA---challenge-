Attribute VB_Name = "Module1"
Sub stock_info():

'currentTicker

'firstOpenPrice = c2
'previousClosedPrice = currentTicker -1 closing price


'stockVolumeSum

'iterate through every row
    'if currentTicker = currentRowAValue
        'stockVolumeSum = stockVolumeSum + column 7
        'previousClosedPrice = cells(i,6).value
    'else
        'print currentTicker
        'print (firstaOpenPrice- previousClosedPrice)
        'print ((firstOpenPrice  - previousClosedPrice)/firstOpenPrice)
        'print (tockVolumeSum)
'run on everyworksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

currentTicker = ws.Cells(2, 1).Value
firstOpenPrice = ws.Cells(2, 3).Value
previousClosedPrice = ws.Cells(2, 6).Value
stockVolumeSum = ws.Cells(2, 7).Value
biggestDub = 0
biggestDubVal = 0
biggestL = 0
biggestLVal = 0
mostStocks = 0
mostStocksVal = 0
numTickers = 1
'assign column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Total Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1

For i = 3 To lRow
    If (currentTicker = ws.Cells(i, 1).Value) Then
        previousClosedPrice = ws.Cells(i, 6).Value
        stockVolumeSum = stockVolumeSum + ws.Cells(i, 7).Value
        
    Else
         
         yearlyChange = previousClosedPrice - firstOpenPrice
         percentChange = yearlyChange / firstOpenPrice
       ' two if loops/conditionals  for highlighting the cells
         If yearlyChange >= 0 Then
                ws.Cells(numTickers + 1, 10).Interior.ColorIndex = 4 ' green
            Else
                ws.Cells(numTickers + 1, 10).Interior.ColorIndex = 3 ' red
            End If
         
          If percentChange >= 0 Then
                ws.Cells(numTickers + 1, 11).Interior.ColorIndex = 4 ' green
            Else
                ws.Cells(numTickers + 1, 11).Interior.ColorIndex = 3 'red
            End If
         
    
       
        If (percentChange > biggestDubVal Or numTickers = 1) Then
        biggestDub = currentTicker
        biggestDubVal = percentChange
        End If
        
        If (percentChange < biggestLVal Or numTickers = 1) Then
        biggestL = currentTicker
        biggestLVal = percentChange
        End If
        
        If (stockVolumeSum > mostStocksVal Or numTickers = 1) Then
        mostStocks = currentTicker
        mostStocksVal = stockVolumeSum
        End If
        
        ws.Cells(numTickers + 1, 9).Value = currentTicker
        ws.Cells(numTickers + 1, 10).Value = yearlyChange
        ws.Cells(numTickers + 1, 11).Value = Format(percentChange, "#.##%")
        ws.Cells(numTickers + 1, 12).Value = stockVolumeSum

         currentTicker = ws.Cells(i, 1).Value
         firstOpenPrice = ws.Cells(i, 3).Value
         previousClosedPrice = ws.Cells(i, 6).Value
         stockVolumeSum = ws.Cells(i, 7).Value
         
         numTickers = numTickers + 1
         
         
    End If

    
Next i

ws.Cells(2, 15).Value = biggestDub
ws.Cells(2, 16).Value = Format(biggestDubVal, "#.##%")
ws.Cells(3, 15).Value = biggestL
ws.Cells(3, 16).Value = Format(biggestLVal, "#.##%")
ws.Cells(4, 15).Value = mostStocks
ws.Cells(4, 16).Value = mostStocksVal

Next ws

End Sub

