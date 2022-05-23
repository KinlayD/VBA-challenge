# VBA-challenge
VBA challenge (homework)

Sub VBA_Summary_Table():

'create summary table with headings
Dim summary_table_headings(4) As String
summary_table_headings(0) = "Ticker"
summary_table_headings(1) = "Yearly Change"
summary_table_headings(2) = "Percent Change"
summary_table_headings(3) = "Total Stock Volume"
Range("I1:L1").Value = summary_table_headings

' summary_table prints yearly_change, percent_change and stock_volume per ticker_name
Dim summary_table As Double
summary_table = 2

' ticker_name holds stock symbol
Dim ticker_name As String

' yearly_change holds absolute yearly change in stock price
Dim yearly_change As Double

' start_row
Dim start_row As Double
start_row = 2

'end_row
Dim end_row As Double
end_row = 2

' percent_change holds relative yearly change in stock price
Dim percent_change As Double
percent_change = 0

' stock_volume holds total stock volume per stock
Dim stock_volume As Double
stock_volume = 0

' assign last_row to the final stock recorded per year
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

' loop through row data
    For i = 2 To Last_Row

' pass loop through an if statement that takes two rows in col A and executes the following statements if condition evaluates to TRUE
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' for each iteration that passes through the ticker column, store cells that are not equal to ticker_name
        ticker_name = Cells(i, 1).Value

' print unique stock symbols to summary_table
            Range("I" & summary_table).Value = ticker_name

' from volume column, store the stock volumes that correspond to ticker_name to the variable stock_volume
        stock_volume = stock_volume + Cells(i, 7).Value

' In summary_table print the stock volume values that correspond to the correct stock symbols
            Range("L" & summary_table).Value = stock_volume

' finding the first row in <open>
                If Cells(start_row, 3).Value = 0 Then

                    For curr_row = start_row To i

                        If Cells(curr_row, 3).Value <> 0 Then

                        start_row = curr_row

                    Exit For
    
                        End If

                Next curr_row
    
                End If
                
                start_row = i + 1

' finding the last row in <close>
                    If Cells(i, 6).Value = 0 Then
    
                    For curr_row_close = end_row To i
                    
                        If Cells(curr_row_close, 6).Value <> 0 Then
                    
                        end_row = curr_row_close
                    
                    Exit For
                    
                        End If
                    
                    Next curr_row_close
    
                    End If
                    
                    end_row = i + 1

' find the difference of last and first row and store to yearly change
yearly_change = Cells(i, 6).Value - Cells(start_row, 3).Value

' print result on summary table
Range("j" & summary_table).Value = yearly_change

' add a new row to summary table
            summary_table = summary_table + 1

' At the end of an iteration reset the counter for stock_volume
            stock_volume = 0

Else

' at each iteration, sum cell values from volume column to stock_volume
stock_volume = stock_volume + Cells(i, 7).Value

End If

Next i

' summary table analysis
Last_Row_Table = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To Last_Row_Table

' find percent range with yearly change and first rows in open
' Range("k" & summary_table).Value = (Cells(i, 6).Value - Cells(start_row, 3).Value) / Cells(start_row, 3).Value * 100 (overflow error)

' conditional statement to filter col 10 values with green if change is positive, red otherwise
    If Cells(j, 10).Value > 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    
    Else
    Cells(j, 10).Interior.ColorIndex = 3

End If

Next j

End Sub

