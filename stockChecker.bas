Attribute VB_Name = "Module1"
Sub stocks()
    For Each ws In Worksheets
        'lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).column don't need this, only want it to go over the data which we know when it ends
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim currentTicker As String
        Dim nextTicker As String
        Dim tickerIncrement As Integer
        tickerIncrement = 2 'keep track of how many tickers we've gone through to put data right, start at 2 for inputting the data
        Dim openYear As Double 'opening value at start of year
        openYear = ws.Cells(2, 3).Value 'set the opening value for the first ticker since it won't be reset until end of the loop
        Dim closeYear As Double 'closing value at end of year, doesn't need to be set yet
        'Dim stockVolume As Long 'long is too small in vba to handle this - for some reason VBA longs are only 4 bytes instead of 8
        Dim stockVolume As Variant 'have to use variant so it can use currency or decimal or something with more bytes
        stockVolume = 0 'volume of stocks for a specific ticker in a year
        Dim yearlyChange As Double 'store change over a year as this just to be easy
        
        'set up all the column headers etc
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To lastRow
            currentTicker = ws.Cells(i, 1).Value
            nextTicker = ws.Cells(i + 1, 1).Value
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            'For j = 2 To lastColumn we don't need lastColumn and starting j at 2 is wrong lmao
            'For j = 1 To 7 'do I even need to iterate through j? I only need specific values not the day to day change
                
            'Next j
            If currentTicker <> nextTicker Then
                closeYear = ws.Cells(i, 6).Value
                yearlyChange = closeYear - openYear
                ws.Cells(tickerIncrement, 9).Value = currentTicker 'ticker name
                ws.Cells(tickerIncrement, 10).Value = yearlyChange 'yearly change
                If ws.Cells(tickerIncrement, 10).Value > 0 Then
                    ws.Cells(tickerIncrement, 10).Interior.ColorIndex = 4 'if a positive change, color green
                ElseIf ws.Cells(tickerIncrement, 10).Value < 0 Then
                    ws.Cells(tickerIncrement, 10).Interior.ColorIndex = 3 'if a negative change, color red
                End If
                If (openYear <> 0) Then 'make sure we aren't dividing by zero just in case
                    ws.Cells(tickerIncrement, 11).Value = FormatPercent(yearlyChange / openYear, 2) 'percent change
                Else
                    ws.Cells(tickerIncrement, 11).Value = "N/A" 'not sure what would be best to put here in this instance so I just put N/A
                End If
                ws.Cells(tickerIncrement, 12).Value = stockVolume 'total stock volume
                tickerIncrement = tickerIncrement + 1 'this will end up being 1 more than the number of total tickers since last one will increment as well
                stockVolume = 0 'reset this so the next ticker can use it
                openYear = ws.Cells(i + 1, 3) 'set the opening value for next year
            End If
        Next i
        
        
        tickerIncrement = tickerIncrement - 1 'put tickerIncrement back to last value
        'MsgBox (ws.Cells(tickerIncrement, 9).Value) 'check tickerIncrement is on the right value
        
        'initialize values to compare against
        'Dim greatestPercentIncrease As Double
        'greatestPercentIncrease = -99999
        'Dim greatestPercentDecrease As Double
        'greatestPercentDecrease = 99999
        'Dim greatestTotalVolume As Double
        'greatestTotalVolume = 0
        
        'actually, let's make these ints keeping track of where they're at to compare against so it's easy to get their tickers without more variables
        Dim greatestPercentIncrease As Integer
        greatestPercentIncrease = 2 'have to start at 2 so it has an actual value, 1 is headers
        Dim greatestPercentDecrease As Integer
        greatestPercentDecrease = 2
        Dim greatestTotalVolume As Integer
        greatestTotalVolume = 2
        For i = 2 To tickerIncrement
            If ws.Cells(i, 11).Value <> "N/A" Then 'don't consider nonapplicable values
                If ws.Cells(i, 11).Value > ws.Cells(greatestPercentIncrease, 11).Value Then
                    greatestPercentIncrease = i
                End If
                If ws.Cells(i, 11).Value < ws.Cells(greatestPercentDecrease, 11).Value Then
                    greatestPercentDecrease = i
                End If
                If ws.Cells(i, 12).Value > ws.Cells(greatestTotalVolume, 12).Value Then
                    greatestTotalVolume = i
                End If
            End If
        Next i
        
        'set each of the values we just found through this loop
        'FormatPercent(yearlyChange / openYear, 2)
        ws.Cells(2, 16).Value = ws.Cells(greatestPercentIncrease, 9).Value
        ws.Cells(2, 17).Value = FormatPercent(ws.Cells(greatestPercentIncrease, 11).Value, 2)
        ws.Cells(3, 16).Value = ws.Cells(greatestPercentDecrease, 9).Value
        ws.Cells(3, 17).Value = FormatPercent(ws.Cells(greatestPercentDecrease, 11).Value, 2)
        ws.Cells(4, 16).Value = ws.Cells(greatestTotalVolume, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(greatestTotalVolume, 12).Value
    Next
End Sub
