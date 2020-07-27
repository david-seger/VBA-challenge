Attribute VB_Name = "Module1"
Sub MultiYearStockData()
'---------------------------------------------------------------------------------
' Amalize stock data over multiple years (worksheets) and summarize by ticker
' the total stock volume and calculate the percent change% and yearly change$
'---------------------------------------------------------------------------------

' Define Fields

Dim lastRow As Long

Dim currentTicker As String
Dim tickerRow As Long

Dim beginPrice As Double
Dim endPrice As Double

Dim tickerVolume As Long
Dim priceChange As Double
Dim percentChange As Double

Dim greatestPercentIncreaseValue As Double
Dim greatestPercentDecreaseValue As Double
Dim greatestTickerVolumeValue As LongLong


' Loop through each worksheet to calculate values

For Each ws In Worksheets
    
    'Cells(Rows,Count, "A").End(x1Up).row will find the last cell with data in column A
    'for the current worksheet
        
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Initialize fields for calculations
    
    currentTicker = ws.Cells(2, 1).Value
    tickerRow = 2
    beginPrice = 0
    endPrice = 0
    tickerVolume = 0
    priceChange = 0
    percentChange = 0
    greatestPercentIncreaseValue = 0
    greatestPercentDecreaseValue = 0
    greatestTickerVolumeValue = 0
    
    'Populate Headers for the new columns and rows in summary section
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("Q2").NumberFormat = "#,##0.00%"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("Q3").NumberFormat = "#,##0.00%"
    ws.Range("O4").Value = "Greatest Total Volume"

    'Start in row 2 (because thats where the stock data starts and
    'in the FOR loop go all the way to the last row
    
    For i = 2 To lastRow
    
        'Check to see if ticker has changed and we have not reached the last row
        
        If currentTicker = ws.Cells(i, 1).Value And i <> lastRow Then
        
            'Get the beginning price - it is the first row of each ticker we know to populate because it is currently zero
            
            If beginPrice = 0 Then
                beginPrice = ws.Cells(i, 3).Value
            End If
            
            'Add to total volume and populate the endPrice that will be used when the ticker changes
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            endPrice = ws.Cells(i, 6).Value
            
        Else
            
            ' Here ticker has changed or we are on the last row with data
            ' do calculations and populate columns
            
            ws.Cells(tickerRow, 9).Value = currentTicker
            priceChange = endPrice - beginPrice
            
            'Begin Price has to be greater than 0 otherwise you will get a divide by zero
            
            If beginPrice > 0 Then
                percentChange = priceChange / beginPrice
            Else
                percentChange = 0
            End If
            
            'Populate ticker data into appropriate cells
            
            ws.Cells(tickerRow, 10).Value = priceChange
            ws.Cells(tickerRow, 11).Value = percentChange
            ws.Cells(tickerRow, 12).Value = totalVolume
    
            'Here we check for the overall greatest increase & decrease and the total volume
            'and if this stock is greater than current values replace.
            
            If percentChange > 0 And percentChange > greatestPercentIncreaseValue Then
                ws.Range("P2").Value = currentTicker
                ws.Range("Q2").Value = percentChange
                greatestPercentIncreaseValue = percentChange
            End If
            
            If percentChange < greatestPercentDecreaseValue Then
                ws.Range("P3").Value = currentTicker
                ws.Range("Q3").Value = percentChange
                greatestPercentDecreaseValue = percentChange
            End If
            
            If totalVolume > greatestTickerVolumeValue Then
                ws.Range("P4").Value = currentTicker
                ws.Range("Q4").Value = totalVolume
                greatestTickerVolumeValue = totalVolume
            End If
    
            'Format the new cells - one is a number format the other a percentage
            
            ws.Cells(tickerRow, 10).NumberFormat = "#,##0.00"
            ws.Cells(tickerRow, 11).NumberFormat = "#,##0.00%"
      
            'Check to see if price has changed for good or bad.  We will set the color index based on that
     
            If ws.Cells(tickerRow, 10).Value >= 0 Then
                    
                ' Positive Change in price is colored green
                
                ws.Cells(tickerRow, 10).Interior.ColorIndex = 4
            
            Else
            
                'Negative Change in price is colored red
                
                ws.Cells(tickerRow, 10).Interior.ColorIndex = 3
                
            End If
     
            ' Reset values to process next ticker
            currentTicker = ws.Cells(i, 1).Value
            beginPrice = ws.Cells(i, 3).Value
            endPrice = ws.Cells(i, 6).Value
            totalVolume = ws.Cells(i, 7).Value
            
            'if not last row add one to advance the ticker row.
            
            If i <> lastRow Then
                tickerRow = tickerRow + 1
            End If
            
        End If
            
    Next i
        
    'Adjust columns to Autofit the data
    
    ws.Columns("L:L").EntireColumn.AutoFit
    ws.Columns("O:O").EntireColumn.AutoFit
    
    'Reset the values for the next worksheet
    
    greatestPercentIncreaseValue = 0
    greatestPercentDecreaseValue = 0
    greatestTickerVolumeValue = 0

Next
    
End Sub


