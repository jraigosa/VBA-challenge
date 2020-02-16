Attribute VB_Name = "Module1"
Sub scrapeStockData()

Dim ws_count As Integer

ws_count = ActiveWorkbook.Worksheets.Count

For j = 1 To ws_count

Worksheets(j).Activate

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim ticker As String
Dim firstRecord As Integer
Dim lastRecord As Integer
Dim volume As Double
Dim firstDay As Integer
Dim lastDay As Integer
Dim numStocks As Integer
Dim stockChange As Double
Dim percentChange As Double
Dim largestIncrease As Double
Dim largestDecrease As Double
Dim largestVolume As Double
Dim largestIncreaseTicker As String
Dim largestDecreaseTicker As String
Dim largestVolumetTicker As String

firstDayRow = 0
lastDayRow = 0

volume = 0
largestIncrease = 0
largestDecrease = 0
largestVolume = 0
numStocks = 0

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(2, 15).Value = "Largest % Increase"
Cells(3, 15).Value = "Largest % Decrease"
Cells(4, 15).Value = "Largest Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For i = 2 To lastRow
    ticker = Cells(i, 1).Value
    currentDay = Cells(i, 2).Value
    volume = volume + Cells(i, 7).Value
    If firstDayRow = 0 Then
        firstDayRow = i
        lastDayRow = i
    ElseIf currentDay < Cells(firstDayRow, 2).Value Then
        firstDayRow = i
    ElseIf currentDay > Cells(lastDayRow, 2).Value Then
        lastDayRow = i
   End If
   If ticker <> Cells(i + 1, 1).Value Then
        numStocks = numStocks + 1
        Cells(numStocks + 1, 9).Value = ticker
        stockChange = Cells(lastDayRow, 6).Value - Cells(firstDayRow, 3).Value
        Cells(numStocks + 1, 10).Value = stockChange
        Cells(numStocks + 1, 10).NumberFormat = "0.00"
        If stockChange > 0 Then
            Cells(numStocks + 1, 10).Interior.ColorIndex = 4
       ElseIf stockChange < 0 Then
            Cells(numStocks + 1, 10).Interior.ColorIndex = 3
       End If
       If Cells(firstDayRow, 3).Value = 0 Then
            percentChange = 0
            Cells(numStocks + 1, 11).Value = Null
       Else
          percentChange = stockChange / Cells(firstDayRow, 3).Value
          Cells(numStocks + 1, 11).Value = percentChange
       End If
        
        Cells(numStocks + 1, 11).NumberFormat = "0.00%"
        Cells(numStocks + 1, 12).Value = volume
        Cells(numStocks + 1, 12).NumberFormat = "0"
        
        If percentChange > largestIncrease Then
            largestIncrease = percentChange
            largestIncreaseTicker = ticker
        ElseIf percentChange < largestDecrease Then
            largestDecrease = percentChange
            largestDecreaseTicker = ticker
        End If
        If volume > largestVolume Then
            largestVolume = volume
            largestVolumeTicker = ticker
        End If
                
        volume = 0
        firstDayRow = 0
        lastDayRow = 0
   End If
            
Next i

Cells(2, 16).Value = largestIncreaseTicker
Cells(3, 16).Value = largestDecreaseTicker
Cells(4, 16).Value = largestVolumeTicker
Cells(2, 17).Value = largestIncrease
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).Value = largestDecrease
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 17).Value = largestVolume

Next j

End Sub

