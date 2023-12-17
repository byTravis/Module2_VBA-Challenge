Attribute VB_Name = "VBA_StockChallenge"
Sub stockData()

For Each ws In Worksheets 'loops through every worksheet

    'establish variables
    Dim lastRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim rowCount As Long
    
    Dim testCount As Integer
    testCount = 2
    

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row  'finds the last row available
    rowCount = 2
    openPrice = 0
    closePrice = 0

    
    ' creates row header titles for breakdown
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("I1:L1").Font.Bold = True
        ws.Range("I:I, J:J, K:K, L:L").Columns.AutoFit

        

        
    

    
    For I = 2 To lastRow   'loops through every row on a specific worksheet
    
        If ws.Cells(I - 1, 1) <> ws.Cells(I, 1) Then 'this is a new record

            openPrice = ws.Cells(I, 3) ' sets open price
            
            ws.Cells(rowCount, 9).Value = ws.Cells(I, 1) ' sets ticker
            ws.Cells(rowCount, 10).Value = 0 ' sets starting yearly change
            ws.Cells(rowCount, 11).Value = 0 ' sets strting percent change
            ws.Cells(rowCount, 12).Value = ws.Cells(I, 7).Value ' sets initial volume
            
            rowCount = rowCount + 1 ' advances row count to the next line for the next record
 
        Else ' updating existing record
            closePrice = ws.Cells(I, 6) ' updates close price
            ws.Cells(rowCount - 1, 10).Value = closePrice - openPrice ' change in price for the year
            ws.Cells(rowCount - 1, 11).Value = ((closePrice - openPrice) / openPrice) ' percent change
            ws.Cells(rowCount - 1, 11).NumberFormat = "0.00%" ' formats number as a percentage
            
            ws.Cells(rowCount - 1, 12).Value = ws.Cells(rowCount - 1, 12).Value + ws.Cells(I, 7) ' updates total volume

        End If
        
        
        
        
        ' conditional formatting for % change
        If ws.Cells(rowCount - 1, 11) > 0 Then  ' positive = green
            ws.Cells(rowCount - 1, 11).Interior.ColorIndex = 4
        
            Else  'negative = red
                ws.Cells(rowCount - 1, 11).Interior.ColorIndex = 3

        End If
        

    Next I
    
    ' Creates and populates summery table
    
        ' creates table for summery report
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ws.Range("O2:O4, O1:Q1").Font.Bold = True
        ws.Range("Q2:Q3").NumberFormat = "0.00%" ' formats number as a percentage
        
        
        'finds Greatest % Increase
        maxVal = Application.WorksheetFunction.Max(ws.Range("K:K"))
        maxRow = Application.WorksheetFunction.Match(maxVal, ws.Range("K:K"), 0)
        ws.Range("P2").Value = ws.Cells(maxRow, 9).Value
        ws.Range("Q2").Value = ws.Cells(maxRow, 11).Value
        
        'finds Greatest % Decrease
        minVal = Application.WorksheetFunction.Min(ws.Range("K:K"))
        minRow = Application.WorksheetFunction.Match(minVal, ws.Range("K:K"), 0)
        ws.Range("P3").Value = ws.Cells(minRow, 9).Value
        ws.Range("Q3").Value = ws.Cells(minRow, 11).Value
        
        'finds greatest Total volume
        maxVolVal = Application.WorksheetFunction.Max(ws.Range("L:L"))
        maxVolRow = Application.WorksheetFunction.Match(maxVolVal, ws.Range("L:L"), 0)
        ws.Range("P4").Value = ws.Cells(maxVolRow, 9).Value
        ws.Range("Q4").Value = ws.Cells(maxVolRow, 12).Value
    

        
    ' autofits columns so whole value is displayed
    ws.Range("I:Q").Columns.AutoFit
    
    



Next ws  ' goes to next worksheet
    

End Sub



