Attribute VB_Name = "stock_data_v1_3"
'Create loop for all stocks and place the following in a summary table
'Ticker symbol I or 9
'Yearly change between first opening price to last closing price 10 or J
'Percent change (Yearly change / opening price) 11 or K
    'Conditional formatting that highlights negative/positive change in green/red
    'Style values with %
    'Total stock volume
'MUST run on every ws and running script once

'BONUS!! Max % increase, Min % decrease, Max total volume
'-------------------------------------------------------------------------------

Sub stock_data():
    
    Dim ws As Integer
    For ws = 1 To Worksheets.Count
    Worksheets(ws).Select
    
    'Define variables
    Range("I1").Value = "Ticker Name"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Dim ticker_name As String, total_volume As LongLong, open_price As Double, close_price As Double, yearly_change As Double, percent_change As Double
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ticker_name = " "
    total_volume = 0
    summary_table = 2
    open_price = Cells(2, 3).Value
    'close_price = Cells(2, 6).Value
            
    'Define Bonus Variables
    
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Max Total Volume"
        Range("O1").Value = "Value"
        Range("P1").Value = "Ticker Name"
    
        Dim max_percent As Double, min_percent As Double, max_volume As LongLong, tickername_index As Integer

    '1) Ticker Symbol
    For Row = 2 To lastrow
        
        If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
            ticker_name = Cells(Row, 1).Value
            Cells(summary_table, 9).Value = ticker_name
            
    '2) Total Stock Volume
            
            total_volume = total_volume + Cells(Row, 7).Value
            Cells(summary_table, 12).Value = total_volume
            total_volume = 0
            
    '3) Yearly Change = Closing price - Opening price
    
            close_price = Cells(Row, 6).Value
            yearly_change = close_price - open_price
            Cells(summary_table, 10).Value = yearly_change
            Cells(summary_table, 10).NumberFormat = "$0.00"
            
    '4) Percent Change = Yearly Change / Opening price
    'Format style as %
    'Conditional Format Yearly Change: Red/Green for Negative/Positive Changes
    
            percent_change = yearly_change / open_price
            Cells(summary_table, 11).Value = percent_change
            Cells(summary_table, 11).NumberFormat = "0.00%"
            
                If Cells(summary_table, 10).Value > 0 Then
                    Cells(summary_table, 10).Interior.ColorIndex = 4
                ElseIf Cells(summary_table, 10).Value < 0 Then
                    Cells(summary_table, 10).Interior.ColorIndex = 3
                Else
                    'Cells(summary_table, 10).Interior.ColorIndex = 0
                End If
                
            open_price = Cells(Row + 1, 3).Value
            summary_table = summary_table + 1
        Else
            total_volume = total_volume + Cells(Row, 7).Value
            
        End If
    Next Row
    
    '5) Bonus Max % Increase, Min % Decrease, Max Total_volume
        
        max_percent = WorksheetFunction.Max(Range("K2:K" & lastrow))
        tickername_index = WorksheetFunction.Match(max_percent, Range("K2:K" & lastrow), 0)
        Range("O2").Value = max_percent
        Range("P2").Value = Range("I" & tickername_index + 1).Value
        
        min_percent = WorksheetFunction.Min(Range("K2:K" & lastrow))
        tickername_index = WorksheetFunction.Match(min_percent, Range("K2:K" & lastrow), 0)
        Range("O3").Value = min_percent
        Range("P3").Value = Range("I" & tickername_index + 1).Value
        
        max_volume = WorksheetFunction.Max(Range("L2:L" & lastrow))
        tickername_index = WorksheetFunction.Match(max_volume, Range("L2:L" & lastrow), 0)
        Range("P4").Value = Range("I" & tickername_index + 1).Value
        Range("O4").Value = max_volume
        
        
        Range("O2:O3").NumberFormat = "0.00%"
        'Range("O4").NumberFormat = "0000 E+0"
        Range("O4").NumberFormat = "##0.00 E+0"
    Next ws
    
End Sub
