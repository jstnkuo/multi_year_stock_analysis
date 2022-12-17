Attribute VB_Name = "Module1"
Sub stock_analysis_loop()

    'loop through all sheets
    For Each ws In Worksheets
        
        Dim WorksheetName As String
        Dim ticker As String
        Dim opening_price, closing_price, yearly_change, percent_change As Double
        Dim stock_total
        stock_total = 0
        Dim greatest_increase, greatest_decrease, greatest_stockv As Double
        Dim i, j As Integer
        i = 0
        j = 0
        greatest_increase = 0
        greatest_decrease = 1
        greatest_stockv = 0
    
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        WorksheetName = ws.Name
    
        For Row = 2 To RowCount
            'This if occurs at the first row of a stock
            If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
                    
                'opening price is the first row column 3
                opening_price = ws.Cells(Row, 3).Value

                'set column I to the ticker
                ws.Range("I" & 2 + j).Value = ws.Cells(Row, 1).Value
                j = j + 1
                                   
            End If
        
            'This if occurs at the last row of a stock
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
                'closing price is the last row of column 3
                closing_price = ws.Cells(Row, 6).Value
            
                'grab stock volume
                stock_total = stock_total + ws.Cells(Row, 7).Value
            
                'grab yearly change and populate
                yearly_change = closing_price - opening_price
                ws.Range("J" & 2 + i).Value = yearly_change
            
                'grab percentage change and populate
                percent_change = (yearly_change / opening_price)
                ws.Range("K" & 2 + i).Value = percent_change
            
                'populate stock total
                ws.Range("L" & 2 + i).Value = stock_total
                i = i + 1
            
                'reset stock total
                stock_total = 0
            Else
                stock_total = stock_total + ws.Cells(Row, 7).Value
            
            End If
        Next Row
    
        For Row = 2 To RowCount
    
            'greatest % increase
            If ws.Cells(Row, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(Row, 11)
                ws.Range("Q" & 2).Value = greatest_increase
            End If
        
            'greatest % decrease
            If ws.Cells(Row, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(Row, 11)
                ws.Range("Q" & 3).Value = greatest_decrease
            End If
        
            'greatest total stock volume
            If ws.Cells(Row, 12).Value > greatest_stockv Then
                greatest_stockv = ws.Cells(Row, 12)
                ws.Range("Q" & 4).Value = greatest_stockv
            End If
        
            'populate stock with greatest increase
            If ws.Cells(Row, 11).Value = greatest_increase Then
                ws.Range("P" & 2).Value = ws.Cells(Row, 9).Value
            End If
        
            'populate stock with greatest decrease
            If ws.Cells(Row, 11).Value = greatest_decrease Then
                ws.Range("P" & 3).Value = ws.Cells(Row, 9).Value
            End If
        
            'populate stock with greatest stock volume
            If ws.Cells(Row, 12).Value = greatest_stockv Then
                ws.Range("P" & 4).Value = ws.Cells(Row, 9).Value
            End If
        
            'conditional color formatting for yearly change
            If ws.Cells(Row, 10) < 0 Then
                ws.Cells(Row, 10).Interior.ColorIndex = 3
            ElseIf Cells(Row, 10) > 0 Then
                ws.Cells(Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Row, 10).Interior.ColorIndex = 0
            End If
        
            'conditional color formatting for percentage change
            If ws.Cells(Row, 11) < 0 Then
                ws.Cells(Row, 11).Interior.ColorIndex = 3
            ElseIf Cells(Row, 11) > 0 Then
                ws.Cells(Row, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(Row, 11).Interior.ColorIndex = 0
            End If

            ws.Cells(Row, 11).Style = "Percent"
            ws.Cells(Row, 11).NumberFormat = "0.00%"
        Next Row
    
        'populate column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'some visual clean up
        ws.Columns("I:I").EntireColumn.AutoFit
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("P:P").EntireColumn.AutoFit
        ws.Columns("Q:Q").EntireColumn.AutoFit
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    Next ws
End Sub
