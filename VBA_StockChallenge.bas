Attribute VB_Name = "VBA_StockChallenge"
Sub stockData()

For Each ws In Worksheets 'loops through every worksheet





    'establish variables
    Dim lastRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim rowCount As Long
    
    
    
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row  'finds the last row available
    rowCount = 2
    Debug.Print lastRow
    
        
    'creates summery table for report
        ' table labels
            'ws.Range("j1").Value = "Ticker"
            'ws.Range("k1").Value = "Value"
            'ws.Range("i2").Value = "Greatest % Increase"
            'ws.Range("i3").Value = "Greatest % Decrease"
            'ws.Range("i4").Value = "Greatest Total Volume"
        ' table formatting
            'ws.Range("I:I").Columns.AutoFit
            'ws.Range("I2:I4, J1:K1").Font.Bold = True
            'ws.Range("I2:I4, J1:K1").Font.Color = vbBlue
    
    
    

    
    
    For i = 2 To lastRow   'loops through every row on a specific worksheet
    
        If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then 'this is a new record
            
            ws.Cells(rowCount, 9).Value = ws.Cells(i, 1)
            
            openPrice = ws.Cells(i, 3)
            ws.Cells(rowCount, 10).Value = openPrice
            ws.Cells(rowCount, 14).Value = ws.Cells(i, 7).Value ' sets initial volume
            
            
            
            rowCount = rowCount + 1
            
            
        
        
        Else ' updating existing record
            closePrice = ws.Cells(i, 6)
            ws.Cells(rowCount - 1, 11).Value = closePrice
            ws.Cells(rowCount - 1, 12).Value = closePrice - openPrice
            ws.Cells(rowCount - 1, 13).Value = ((closePrice - openPrice) / openPrice) ' percent change
            'Format(s.Cells(rowCount - 1, 13).Value, "0.000%")
            
            ws.Cells(rowCount - 1, 14).Value = ws.Cells(rowCount - 1, 14).Value + ws.Cells(i, 7) ' updates total volume
            
            
            
        
        End If
        
    
    
    Next i
    
        
        
   

        
        
    
    



Next ws  ' goes to next worksheet
    

End Sub



