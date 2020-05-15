Sub Stocks()
'Variable definitions:
    'row_i is the current empty row in column I where unique tickers will be entered
    Dim row_i As Integer
    'last_row is the last populated row in column A.
    Dim last_row As LongLong
    'ticker is the stock ticker value in a given row in column A
    Dim ticker As String
    'row is the current row in a loop
    Dim row As LongLong
    'opening price is the opening price value (column C) of a stock
    Dim opening_price As Double
    'closing price is the closing price value (column F)
    Dim closing_price As Double
    'change in price over the course of the year:
    Dim price_change As Double
    'define percent change
    Dim percent_change As Double
    'Define total stock volume
    Dim stock_vol As LongLong
    'Define i for array counts
    Dim i As Integer
    'define headers for arrays
    Dim headers() As Variant
    'define inputs for arrays
    Dim inputs() As Variant
      
    
'PART 1......
'(Challenge Part 2) Have code work across all worksheets (reference:https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook)
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Select
    
     '1. Add headers to summary table from an array
    headers() = Array("Ticker", "Yearly Change ($)", "Percent Change", "Total Stock Volume")
     For i = 0 To 3
         Cells(1, 9 + i).Value = headers(i)
     Next i
     
     '2. Initiate row counter for columns i - k inputs at row 2:
     row_i = 2
     
     '3. Iniate opening price as first value for first ticker
     opening_price = (Cells(2, 3).Value)
     '   check value of opening price
         'MsgBox (opening_price)
        
     '4. Determine how many rows of information are in the source. Define this value as last_row
     last_row = Cells(Rows.Count, 1).End(xlUp).row
     
     
     '5. Initialize variable stock_vol to zero
     stock_vol = 0
     
     '6. Loop through all the stocks for one year
     For row = 2 To (last_row)
     
     '7. Add this row's stock volume to existing stock volume
         stock_vol = stock_vol + Cells(row, 7).Value
         'MsgBox (stock_vol)
       
         '9. If ticker symbol in selected row and column A is not the same as the value in the same column and next row then
         If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
         
             '9.1. Select the ticker value before the change and enter this value into row_i, column I.
              ticker = Cells(row, 1).Value
              'MsgBox (ticker)
              
             '9.2. Define closing price
             closing_price = (Cells(row, 6).Value)
             'MsgBox (closing_price)
             
             '9.3. Define yearly change in price as opening price - closing price
             price_change = FormatCurrency(opening_price - closing_price)
            ' MsgBox (price_change)
             
             '9.4.  Define percent change as the price_change/ opening_price, output it in column j & format at % (Ref. https://www.excelfunctions.net/vba-formatpercent-function.html)
             'when looping across sheets found that the next line of code would error out if the opening price was zero so created a condition to set percent change equal to zero in that case.
             
            If opening_price <> 0 Then
                percent_change = (price_change / opening_price)
            Else: percent_change = 0
            End If
            
             '9.5. create and output array of values for ticker, yearly change in price, percent change and stock volume
             inputs() = Array(ticker, price_change, percent_change, stock_vol)
             For i = 0 To 3
                 Cells(row_i, 9 + i).Value = inputs(i)
               
             Next i

                     
             '9.6. format the cells in column 10 (price_change) so that if negative they are red, if positive they are green
             If Cells(row_i, 10).Value < 0 Then
                 Cells(row_i, 10).Interior.ColorIndex = 3
                 Else: Cells(row_i, 10).Interior.ColorIndex = 4
                 
             End If
             
             '9.7. Reset initial stock volume for ticker to zero
             stock_vol = 0
                     
             '9.8. Increase row_i count by one.
             row_i = row_i + 1
             
             '9.9 reset opening price to next row+1
             opening_price = (Cells(row + 1, 3))
            ' MsgBox (opening_price)
              
         End If
         
     '10. Continue looping through rows.
     Next row
      'to fix formatting issues in percentage after a zero value:
                Range(Cells(2, 11), Cells(last_row, 11)).Select
                Selection.Style = "Percent"
                Selection.NumberFormat = "0.00%"
     
     'Challenge 1:
     ' Determine and display greatest percent increase, greatest percent decrease and greatest total volume.
     
     'Define additional variables
      Dim greatest_percent_inc As Double
      Dim greatest_percent_dec As Double
      Dim greatest_total_vol As LongLong
      Dim headers2() As Variant
      Dim ticker_gpi As String
      Dim ticket_gpd As String
      Dim ticker_tsv As String
      Dim outputs2() As Variant
      Dim outputs3() As Variant
      Dim j As Integer
      
      
      '1. Create headers for final calculated fields
      headers2() = Array("Greatest Percent Increase", "Greatest Percent Decrease", "Greatest Total Volume")
      'inputs consecutive array values from headers2() into cells N2, N3 and N4.
      For i = 0 To 2
         Range("N" & i + 2).Value = headers2(i)
      Next i
      
      'defines headers for cells O1 and P1
      Range("O1").Value = "Ticker"
      Range("P1").Value = "Value"
      
      '2.initiate values
      greatest_percent_inc = 0
      greatest_percent_dec = 0
      greatest_total_vol = 0
     
      
     '3.Define last_row for the annual totals
      last_row = Cells(Rows.Count, 9).End(xlUp).row
      
      '4. Loop through annual totals checking each time if the associated column values are greater than (for greatest percent increase and greater total vol)
      'or less than (for greatest percent decrease), the previous value.  If so than re-define value as that row's value.
      For row = 2 To last_row
         'compare current value for greatest percent increase to the one in this row, if the value in this row is greater than exsiting value record it as the new greatest percent increase.
         If Cells(row, 11).Value > greatest_percent_inc Then
         greatest_percent_inc = Cells(row, 11).Value
         ticker_gpi = Cells(row, 9).Value
         End If
         
         'compare current value for greatest percent decrease to the one in this row, if the value in this row is less than existing value, record it as the new greatest percent increase.
         If Cells(row, 11).Value < greatest_percent_dec Then
         greatest_percent_dec = Cells(row, 11).Value
         ticker_gpd = Cells(row, 9).Value
         End If
         
         'compare current value for greatest stock volume to the one in this row, if the value in this row is greater than the existing value, record it as the new greatest stock volume.
         If Cells(row, 12).Value > greatest_total_vol Then
         greatest_total_vol = Cells(row, 12).Value
         ticker_tsv = Cells(row, 9).Value
         End If
         
      Next row
     
      '5. output values for greatest percent increase, greatest percent decrease and greatest total volume into summary table.
        outputs2() = Array(ticker_gpi, ticker_gpd, ticker_tsv)
        outputs3() = Array(FormatPercent(greatest_percent_inc), FormatPercent(greatest_percent_dec), greatest_total_vol)
        For i = 0 To 2
             Range("O" & i + 2).Value = outputs2(i)
             Range("P" & i + 2).Value = outputs3(i)
         Next i
     
     'check the name of the worksheet
     'MsgBox (ws.Name)
   
Next