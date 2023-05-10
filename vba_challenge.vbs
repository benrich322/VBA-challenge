Sub VBA_Challenge()
' Declare Current as a worksheet object variable.
Dim Current As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each Current In Worksheets

    ' Assign Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' Assign the start row for the summary table
    summary_table_row = 2
    
    ' Identifies last row in the table
     last_row = Cells(Rows.count, 1).End(xlUp).Row
     
    ' Initial value for the row count
     count = 0
     
    ' Loop through all stocks
    For I = 2 To last_row

        'Identifies the value of the volume column for each row
        total_stock_volume = Cells(I, 7).Value
        
        'Adds the row volume to the total ticker volume
        sum_total_stock_volume = sum_total_stock_volume + total_stock_volume
        
        ' Check if we are still within the same stock ticker, if it is not...
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            ' Set the ticker name
            ticker = Cells(I, 1).Value
        
            'Assign the first and last ticker row
            first_ticker_row = Cells(I - count, 3).Value
            last_ticker_row = Cells(I, 6).Value
            
            'Calculates the difference from the opening price at the beginning of the year to the closing price at the end of that year
            yearly_change = last_ticker_row - first_ticker_row
            
            'Calculates the percent change of the opening price at the beginning of a given year to the closing price at the end of that year
            percent_change = (last_ticker_row - first_ticker_row) / first_ticker_row
        
            ' Print the ticker value in the summary table
            Range("I" & summary_table_row).Value = ticker
            
            ' If yearly change is greater than 0 then highlight green if not highlight red
            If yearly_change > 0 Then
                Range("J" & summary_table_row).Value = yearly_change
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_table_row).Value = yearly_change
                Range("J" & summary_table_row).Interior.ColorIndex = 3
                
            End If
            
            ' Format the percent change to a percent
            Range("K" & summary_table_row).Value = FormatPercent(percent_change)
            
            ' Print the sum of the total stock volume in summary table
            Range("L" & summary_table_row).Value = sum_total_stock_volume
        
            ' Add one to the summary table row
            summary_table_row = summary_table_row + 1
            ' Reset the count to 0
            count = 0
            ' Reset the sum of the total stock volume to 0
            sum_total_stock_volume = 0
            
        ' If the cell immediately following a row is the same ticker...
        Else
            ' Increase the count by 1
            count = count + 1
           
        End If
        
    Next I
    
    ' Assign headers to min max summary table
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    ' Identifies last row in summary table
     last_summary_row = Cells(Rows.count, 9).End(xlUp).Row
    
    ' Assigns the first values for the min max summary table
    highest = Cells(2, 11).Value
    highest_ticker = Cells(2, 9).Value
    lowest = Cells(2, 11).Value
    lowest_ticker = Cells(2, 9).Value
    highest_volume = Cells(2, 12).Value
    highest_volume_ticker = Cells(2, 9).Value
    
     
    ' Loop through percent change
    For j = 2 To last_summary_row
    
        ' Check if next value is greater than previous max
        If Cells(j, 11).Value > highest Then
            highest = Cells(j, 11).Value
            highest_ticker = Cells(j, 9).Value
            
        ' Check if next value is less than previous min
        ElseIf Cells(j, 11).Value < lowest Then
            lowest = Cells(j, 11).Value
            lowest_ticker = Cells(j, 9).Value
        
        ' Check if next value is greater than previous max
        ElseIf Cells(j, 12).Value > highest_volume Then
            highest_volume = Cells(j, 12).Value
            highest_volume_ticker = Cells(j, 9).Value
          
        End If
        
    Next j
    
    ' Print the greatest increase, greatest decrease, and greatest total volume
    Cells(2, 17).Value = FormatPercent(highest)
    Cells(2, 16).Value = highest_ticker
    Cells(3, 17).Value = FormatPercent(lowest)
    Cells(3, 16).Value = lowest_ticker
    Cells(4, 17).Value = highest_volume
    Cells(4, 16).Value = highest_volume_ticker

Next Current

End Sub
