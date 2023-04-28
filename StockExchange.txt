Attribute VB_Name = "Module3"
'Create the subroutine
Sub StockMarket3()
    
    'Declare variables
    Dim ws As Worksheet         'For our initial for loop so it loops through all worksheets
    Dim lastrow As Long         'last row of the sheet
    Dim year_start As Double    'stock's yearly initial value
    Dim year_end As Double      'stock's yearly final value
    Dim year_diff As Double     'stock's differnce between initial and final value
    Dim year_diff_pct As Double 'percent growth/decline over the year for a stock based on initial value
    Dim ticker_count As Integer 'Used to place info in rightmost cells
    Dim volume As LongLong      'volume of each stock
    Dim MyRange As Range        'Used to designate ranges of formatting cells after for loop
    Dim gincrease As Double     'Hold value of the greatest stock increase
    Dim gdecrease As Double     'Hold value of the greatest stock decrease
    Dim gvolume As LongLong     'Hold value of the greatest stock volume
    Dim iTicker As String       'To store the ticker of greatest increase
    Dim dTicker As String       'To store the ticker of greatest decrease
    Dim vTicker As String       'To store the ticker of greatest volume
    
    'We want all the below to be reset for each worksheet, so we start our for loop for the worksheets here (googled this bit of code, got from stackoverflow)
    For Each ws In ThisWorkbook.Worksheets
    
        'Initial values
        ticker_count = 2                    'since we want the first ticker in the second row of the sheet
        year_start = ws.Cells(2, 3).Value   'Our first year_start value
        volume = 0                          'volume begins at 0
        
        'Formula for lastrow *code from class slides/stack overflow*
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'NEW: For the headers in columns I through L we want certain values
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Yearly Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        
        'Initialize the for loop
        For i = 2 To lastrow
            
            'Determine where one ticker ends and the next begins
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'Add the last stocks to the volume
                volume = volume + ws.Cells(i, 7).Value
                
                'Capture the year_end value
                year_end = ws.Cells(i, 6).Value
                
                'Find the yearly difference
                year_diff = year_end - year_start
                
                'Find the yearly percent difference
                year_diff_pct = Round(year_diff / year_start, 4)
                
                'Print out the values in the right of the worksheet
                ws.Cells(ticker_count, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(ticker_count, 10).Value = year_diff
                ws.Cells(ticker_count, 11).Value = year_diff_pct
                ws.Cells(ticker_count, 12).Value = volume
                
                'increase ticker_count so as to list all tickers
                ticker_count = ticker_count + 1
                
                'Restart the volume counter
                volume = 0
           
                'Replace the old year_start value with the new ticker's year_start value
                year_start = ws.Cells(i + 1, 3).Value
                
            'Else statement (as in, if the two are the same ticker, just continue to add the volume)
            Else
                
                'Add number of stocks sold to volume
                volume = volume + ws.Cells(i, 7).Value
            
            
            End If
            
        Next i
    
        
        
        'Note that we need to set the last row of the ticker column to a different value than the last row of the main columns
        lastrow_ticker = ws.Cells(Rows.Count, 10).End(xlUp).Row        'returns the last row in column J, which will be the same for columns K and L
        
        'Now we can format the J and K columns
        'Next we want the J column to return red if negative, gray if 0 and green if positive
        For j = 2 To lastrow_ticker
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4 '4 = green
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3 '3 = red
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 15 '15 = grey
                
            End If
        Next j
        
        'Also we want the K column to be formatted for percentages (code from question on stackoverflow)
        ws.Range("K2:K" & lastrow_ticker).NumberFormat = "0.00%"
        
        'Finally we want the greatest percent increase, percent decrease, and total volume of the stocks
        'For this we use variables to store the greatest of each increase, decrease, and volume
        gincrease = 0
        gdecrease = 0
        gvolume = 0
        iTicker = ""
        dTicker = ""
        vTicker = ""
        
        'Now we need a for loop to determine the greatest of these, and then paste them into the cell values at the end
        For k = 2 To lastrow_ticker
        
        
            If ws.Cells(k, 11).Value > gincrease Then
                gincrease = ws.Cells(k, 11).Value
                iTicker = ws.Cells(k, 9).Value
            End If
            If ws.Cells(k, 11).Value < gdecrease Then
                gdecrease = ws.Cells(k, 11).Value
                dTicker = ws.Cells(k, 9).Value
            End If
            If ws.Cells(k, 12).Value > gvolume Then
                gvolume = ws.Cells(k, 12).Value
                vTicker = ws.Cells(k, 9).Value
            End If
        Next k
        
        'Label rows of our greatest increase decrease and volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Label columns of said rows
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'Input the correct ticker in the correct place
        ws.Cells(2, 15).Value = iTicker
        ws.Cells(3, 15).Value = dTicker
        ws.Cells(4, 15).Value = vTicker
        
        'Input the values in the correct places
        ws.Cells(2, 16).Value = gincrease
        ws.Cells(3, 16).Value = gdecrease
        ws.Cells(4, 16).Value = gvolume
        
        'Format the cells the same as the main columns
        ws.Range("P2:P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "0,000"
        
        'For formatting we want the column width to be correct (Note: code sourced from https://powerspreadsheets.com/excel-vba-column-width/#4-AutoFit-Column-Width-Based-on-Entire-Column )
        ws.Range("A1:P1").EntireColumn.AutoFit
        
    Next ws
    
End Sub
