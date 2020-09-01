Attribute VB_Name = "Module1"
Sub stock_market()

'Set dimensions
    Dim ws As Integer
    Dim ticker As String
    Dim tickers_count As Integer
    Dim last_row As Long
    Dim open_price As Double
    Dim closing_price As Double
    Dim annual_price_change As Double
    Dim percent_change As Double
    Dim volume As Double
    
'Loop over each worksheet within the workbook
For ws = 1 To Sheets.Count
    Sheets(ws).Activate

'Add headers for each worksheet
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Determine the last row of the dataset
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
'set initial values for various variables to 0
    tickers_count = 0
    annual_price_change = 0
    open_price = 0
    percent_change = 0
    volume = 0
   
'loop through list of tickers
For i = 2 To last_row

    'Retrieve value of ticker symbol
    ticker = Cells(i, 1).Value
    
    'Retrieve start of year opening price of ticker
    If open_price = 0 Then
        open_price = Cells(i, 3).Value
    End If
    
    'Add total stock volume for ticker
    volume = volume + Cells(i, 7).Value
    
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'increase number of tickers when loop reaches a different ticker in the list
    tickers_count = tickers_count + 1
    
    'print into destination
    Cells(tickers_count + 1, 9) = Cells(i, 1).Value
    
    'retrieve end of year closing price for ticker
    closing_price = Cells(i, 6).Value
    
    'get price change for the year
    annual_price_change = closing_price - open_price
    
    'print price change into destination
    Cells(tickers_count + 1, 10).Value = annual_price_change
    
    'inserting colour codes
    If annual_price_change > 0 Then
        Cells(tickers_count + 1, 10).Interior.ColorIndex = 4
        
    ElseIf annual_price_change < 0 Then
        Cells(tickers_count + 1, 10).Interior.ColorIndex = 3
        
    Else
        Cells(tickers_count + 1, 10).Interior.ColorIndex = 0
    End If
    
    'convert price change into percentage
    If open_price = 0 Then
        percent_change = 0
    
    Else
    
        percent_change = (annual_price_change / open_price)
    
    End If
    
    'format percent_change value as a percentage and print into destination
    Cells(tickers_count + 1, 11).Value = Format(percent_change, "percent")
    
    'set open price back to 0 for different ticker
    open_price = 0
    
    'add total stock volume to destination
    Cells(tickers_count + 1, 12).Value = volume
    
    'set volume back to 0
    volume = 0
    
    End If

Next i


'challenges

'set dimensions
Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_total_volume As Double

'insert header titles
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
'determining the last row of the dataset
    last_row = Cells(Rows.Count, 9).End(xlUp).Row
    
'setting initial values of various variables
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_total_volume = Cells(2, 12).Value
    greatest_total_volume_ticker = Cells(2, 9).Value
    
'loop through array of values
    For i = 2 To last_row
    
    'set conditions
    If Cells(i, 11).Value > greatest_percent_increase Then
        greatest_percent_increase = Cells(i, 11).Value
        greatest_percent_increase_ticker = Cells(i, 9).Value
    End If
    
    If Cells(i, 11).Value < greatest_percent_decrease Then
        greatest_percent_decrease = Cells(i, 11).Value
        greatest_percent_decrease_ticker = Cells(i, 9).Value
    End If
        
    If Cells(i, 12).Value > greatest_total_volume Then
        greatest_total_volume = Cells(i, 12).Value
        greatest_total_volume_ticker = Cells(i, 9).Value
    End If
    
    Next i
    
'print results into the table
    Cells(2, 15).Value = greatest_percent_increase_ticker
    Cells(2, 16).Value = Format(greatest_percent_increase, "percent")
    Cells(3, 15).Value = greatest_percent_decrease_ticker
    Cells(3, 16).Value = Format(greatest_percent_decrease, "percent")
    Cells(4, 15).Value = greatest_total_volume_ticker
    Cells(4, 16).Value = greatest_total_volume
    
Next ws

End Sub

