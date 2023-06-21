Attribute VB_Name = "Module1"
'Module 2 Challenge: VBA-challenge
'Author: Daiana Spataru
'Date: June 19, 2023

'The stock_assess function loops through all the stocks for one year and outputs the following information:
'1. Ticker symbol.
'2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'4. The total stock volume of the stock.

'Sub stock_assess(ws As Worksheet)
 Sub stock_assess()
 
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        'variable declaration
        Dim ticker As String
        Dim numTicker As Integer 'this variable is used to determine the count of each ticker name in the worksheet
        Dim totStockVol As Double 'this variable is used to keep track of the total volume in the summary table
        Dim summTableRow As Integer 'this variable is used to keep track of the location of each ticker in the summary table
        
        'variable assignment
        numTicker = 0
        totStockVol = 0
        summTableRow = 2
    
        numRowsinWS = ws.Cells(Rows.Count, 1).End(xlUp).Row 'calculating the total number of rows (e.g. data) in the worksheet
        
        'adding summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To numRowsinWS
            
            'Conditional statement to check if the next entry is the same as the current entry
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value 'setting the ticker name
                ws.Range("I" & summTableRow).Value = ticker 'printing ticker name in column I
                
                openingPrice = ws.Cells(i - numTicker, 3) 'finding the opening price of a stock
                
                numTicker = 0 'reset ticker count to 0
                
                yearChange = ws.Cells(i, 6).Value - openingPrice 'calculating the yearly change
                ws.Range("J" & summTableRow).Value = yearChange 'printing the yearly change in column J
                
                'changing the colour of the cells in the summary table column
                If ws.Range("J" & summTableRow).Value > 0 Then
                    ws.Range("J" & summTableRow).Interior.ColorIndex = 4 'changing cell colour to green if the yearly change in price of the slock is above 0
                ElseIf ws.Range("J" & summTableRow).Value <= 0 Then
                    ws.Range("J" & summTableRow).Interior.ColorIndex = 3 'changing cell colour to red if the yearly change in price of the slock is below 0
                End If
                
                percentChange = (yearChange / openingPrice) 'calculating the percent change by using the yearChange and dividing it by the opening price, then multiplying by 100
                ws.Range("K" & summTableRow).Value = percentChange 'printing the percent change in column K of the summary table
                
                totStockVol = totStockVol + ws.Cells(i, 7).Value
                ws.Range("L" & summTableRow).Value = totStockVol 'printing the total stock volume in column L of the summary table
                
                summTableRow = summTableRow + 1 'adding one to the summary table row to populate the next entry
                totStockVol = 0 'resetting the total stock volume
                
            'if the next cell has the same ticker name then we execute as below
            Else
                numTicker = numTicker + 1 'keep count of the number of entries in the ticker to determine when the last entry is
                totStockVol = totStockVol + ws.Cells(i, 7).Value 'adding the total stock volume from the previous entry to the current entry
            End If
        
        Next i
        
        'creating a new table to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        'creating labels for the high level summary tables
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'variable declaration
        Dim greatestInc As Double
        Dim greatestDec As Double
        Dim greatestTotVol As Double
        
        'variable assignment
        greatestInc = WorksheetFunction.Max(ws.Range("K:K"))
        greatestDec = WorksheetFunction.Min(ws.Range("K:K"))
        greatestTotVol = WorksheetFunction.Max(ws.Range("L:L"))
        
        'populating the cells with the summary information
        ws.Cells(2, 17).Value = greatestInc
        ws.Cells(3, 17).Value = greatestDec
        ws.Cells(4, 17).Value = greatestTotVol
        

        numRowsinSummTable = ws.Cells(Rows.Count, 11).End(xlUp).Row 'calculating the total number of rows (e.g. data) in the summary table
        
        'it is required to clear the format in order to run the macro on an already formatted sheet otherwise you run into an error with ra.Address
        ws.Range("K2:K" & numRowsinSummTable).ClearFormats
        ws.Range("Q2:Q3").NumberFormat = ClearFormats
        
        'finding the address of each variable to determine the ticker associated with the value
        Dim ra As Range
        
        'finding and populating the ticker for the greatest % increase
        Set ra = ws.Range("K2:K" & numRowsinSummTable).Find(greatestInc)
        Row = Split(ra.Address, "$")(2)
        ws.Cells(2, 16).Value = ws.Cells(Row, 9).Value
        
        'finding and populating the ticker for the greatest % decrease
        Set ra = ws.Range("K2:K" & numRowsinSummTable).Find(greatestDec)
        Row = Split(ra.Address, "$")(2)
        ws.Cells(3, 16).Value = ws.Cells(Row, 9).Value
        
        'finding and populating the ticker for the greatest total volume
        Set ra = ws.Range("L2:L" & numRowsinSummTable).Find(greatestTotVol)
        Row = Split(ra.Address, "$")(2)
        ws.Cells(4, 16).Value = ws.Cells(Row, 9).Value
        
        ws.Columns.AutoFit 'autofit columns to make the data look nice
        ws.Range("K2:K" & numRowsinSummTable).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    Next ws
    
End Sub



