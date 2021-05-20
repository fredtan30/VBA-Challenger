Sub StockTicker()
    
    'Define all variable
    Dim ticker As String
    Dim total_stock_vol As Double
    Dim ticker_counter As Double
    Dim open_counter As Double
    Dim open_market_value As Double
    Dim close_market_value As Double

    Dim last_row As Double
    
    For Each ws In Worksheets
    
    'Set total Stock volume as 0
    total_stock_vol = 0
    
    'set counter to start from row 2 to track rows
    ticker_counter = 2
    open_counter = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 2 To last_row

        total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        open_market_value = ws.Cells(open_counter, 3)
        
        'Comparing the ticker value is different then summarise
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            close_market_value = ws.Cells(i, 6)
            ws.Cells(ticker_counter, 9).Value = ticker  'to display ticker in column "I"
            ws.Cells(ticker_counter, 10).Value = close_market_value - open_market_value
           
           
            If open_market_value = 0 Then
            
                ws.Cells(ticker_counter, 11).Value = Null 'Set wording to null if opening is 0 to avoid the difference divided by 0
                
                
            Else
            
                ws.Cells(ticker_counter, 11).Value = (close_market_value - open_market_value) / open_market_value
                ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
                
                
            End If
            Cells(ticker_counter, 12).Value = total_stock_vol
            
            'Formatting the ticker colour more than 0 is green, less than 0 is red
            If ws.Cells(ticker_counter, 10).Value > 0 Then
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
            Else
               ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
            End If
            
        
            
            ' reset total stock volume to 0
            total_stock_vol = 0
            
        
            'counting the ticker counter for next row
            ticker_counter = ticker_counter + 1
            open_counter = i + 1
        
        End If
        
    Next i
    
    Next ws
    
End Sub