Attribute VB_Name = "Module2"
Sub multiple_sheets2()

    For Each ws In Worksheets

'Set inicial variables

    Dim tickers As String

    Dim stock_volume As LongLong
    
    stock_volume = 0

    Dim Summary_table_row As Long
    
    Dim stock_change As Double
    stock_change = 0
    
    Dim opening_price As Double
    opening_price = ws.Cells(2, 3).Value
    
    Dim closing_price As Double
    
    Dim percentage_change As Double
    
    
'Creating summary table

    Summary_table_row = 2

'find the last row with data

    Dim LR As Long
    LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through stock data

    For i = 2 To LR
    
'Summary table headers

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "percentage change"
    ws.Cells(1, 12).Value = "total stock volume"

        
    ws.Range("I:L").Columns.AutoFit
        

'Getting unique tickers from stock list and sum volume up

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    tickers = ws.Cells(i, 1).Value
    stock_volume = stock_volume + ws.Cells(i, 7).Value
    closing_price = ws.Cells(i, 6).Value
    stock_change = closing_price - opening_price
    
'summary table conditional formatting

    If stock_change > 0 Then
    ws.Cells(Summary_table_row, 10).Interior.ColorIndex = 4
    Else: ws.Cells(Summary_table_row, 10).Interior.ColorIndex = 3
    End If
    
'printing yearly change on summary table
    
    ws.Cells(Summary_table_row, 10).Value = stock_change
    
'printing percentage change on summary table

    percentage_change = (closing_price - opening_price) / opening_price
    ws.Cells(Summary_table_row, 11).Value = percentage_change
    
    ws.Cells(Summary_table_row, 11).NumberFormat = "0.00%"
    
    
'setting opening price for the next stock

    opening_price = ws.Cells(i + 1, 3).Value
    
    
'Printing tickers to the summary table

    ws.Cells(Summary_table_row, 9).Value = tickers

'Printing stock volume to the summary table
    ws.Cells(Summary_table_row, 12).Value = stock_volume

'Add one to summary table row

    Summary_table_row = Summary_table_row + 1

'Reset stock volume

    stock_volume = 0

'If the cell immediately following a row is the same stock
    
    Else

    stock_volume = stock_volume + ws.Cells(i, 7).Value
    
    End If


    Next i
      
    
'printing the greatest values
    
        Dim greatest_increase As Double
            greatest_increase = 0
        
        Dim greatest_decrease As Double
            greatest_decrease = 0
    
        Dim greatest_total As LongLong
            greatest_total = 0
    
        Dim LR_summary As Long
        
        LR_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For j = 1 To LR_summary
    
'printing headers

    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    ws.Cells(2, "O").Value = "greatest % increase"
    ws.Cells(3, "O").Value = "greatest % decrease"
    ws.Cells(4, "O").Value = "greatest total volume"

    ws.Range("O:Q").Columns.AutoFit
    

    If ws.Cells(j + 1, 11).Value > greatest_increase Then
        greatest_increase = ws.Cells(j + 1, 11).Value
        
            ws.Cells(2, 17).Value = greatest_increase
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 16).Value = Cells(j + 1, 9).Value
    
    End If
    
    If ws.Cells(j + 1, 11).Value < greatest_decrease Then
        greatest_decrease = ws.Cells(j + 1, 11).Value
           ws.Cells(3, 17).Value = greatest_decrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(3, 16).Value = ws.Cells(j + 1, 9).Value
    
    End If


    If ws.Cells(j + 1, 12).Value > greatest_total Then
        greatest_total = ws.Cells(j + 1, 12).Value
            ws.Cells(4, 17).Value = greatest_total
            ws.Cells(4, 16).Value = ws.Cells(j + 1, 9).Value
    
    End If
    
        Next j


    Next ws
       

End Sub
