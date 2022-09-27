Attribute VB_Name = "Module1"
Sub StockSummary()

    Dim ws As Worksheet
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    
'ThisWorkbook always returns the workbook in which the code is running.
    For Each ws In ThisWorkbook.Worksheets
    
'set /column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
     
'setup integers for loop
    Summary_Table_Row = 2
    op_row = 2
    
'loop
    For i = 2 To ws.UsedRange.Rows.Count
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
        'find all the values
            Ticker = ws.Cells(i, 1).Value
            
'vol = ws.Cells(i, 7).Value
            year_open = ws.Cells(op_row, 3).Value
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open
            total_volume = total_volume + ws.Cells(i, 7).Value
                
        'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = Ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
            ws.Cells(Summary_Table_Row, 12).Value = total_volume
            
If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
    ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 255, 0)
 ElseIf ws.Cells(Summary_Table_Row, 10).Value < 0 Then
   ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(255, 0, 0)

 End If
                  
            Summary_Table_Row = Summary_Table_Row + 1
        
            total_volume = 0
            op_row = i + 1
            
            Else
            total_volume = total_volume + ws.Cells(i, 7).Value

                 
            End If
        Next i

'finish loop
    
Next
End Sub
