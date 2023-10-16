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
NumTickers = 1
'assign column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1

For i = 3 To lRow
    If (currentTicker = ws.Cells(i, 1).Value) Then
        previousClosedPrice = ws.Cells(i, 6).Value
        stockVolumeSum = stockVolumeSum + ws.Cells(i, 7).Value
        
    Else
         
         yearlyChange = firstOpenPrice - previousClosedPrice
         percentChange = yearlyChange / firstOpenPrice
       ' two if loops/conditionals  for highlighting the cells
         If yearlyChange >= 0 Then
                ws.Cells(NumTickers + 1, 10).Interior.ColorIndex = 4 ' green
            Else
                ws.Cells(NumTickers + 1, 10).Interior.ColorIndex = 3 ' red
            End If
         
          If percentChange >= 0 Then
                ws.Cells(NumTickers + 1, 11).Interior.ColorIndex = 4 ' green
            Else
                ws.Cells(NumTickers + 1, 11).Interior.ColorIndex = 3 'red
            End If
         
            ws.Cells(NumTickers + 1, 9).Value = currentTicker
            ws.Cells(NumTickers + 1, 10).Value = yearlyChange
            ws.Cells(NumTickers + 1, 11).Value = Format(percentChange, "#.##%")
            ws.Cells(NumTickers + 1, 12).Value = stockVolumeSum

         currentTicker = ws.Cells(i, 1).Value
         firstOpenPrice = ws.Cells(i, 3).Value
         previousClosedPrice = ws.Cells(i, 6).Value
         stockVolumeSum = ws.Cells(i, 7).Value
         
         NumTickers = NumTickers + 1
         
         
    End If

    
Next i


Next ws

End Sub

