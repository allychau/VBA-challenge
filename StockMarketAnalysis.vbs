Attribute VB_Name = "Module1"
'VBA Homework - The VBA of Wall Street
Sub StockMarketAnalysis()

     For Each ws In Worksheets
     
        'Label new column headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        Row = 2
        Column = 1
       
        TotalTickerVolumn = 0
        
        ' Get the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (lastRow)
        
        'Set initial open price
        openPrice = ws.Cells(2, Column + 2).Value
        
        For i = 2 To lastRow
            
            'Check if the next Ticker Symbol is the same as the next Ticker
            If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then
            
                'Set Ticker Name
                tickerName = ws.Cells(i, Column).Value
                ws.Cells(Row, Column + 8).Value = tickerName
                 
                 closePrice = ws.Cells(i, Column + 5).Value
                 
                 'Add Yearly Price
                 yearlyChange = closePrice - openPrice
                 ws.Cells(Row, Column + 9).Value = yearlyChange
                 
                 'Add Percent Change
                 If openPrice = 0 Then
                    percentChange = 0
                 Else
                     percentChange = yearlyChange / openPrice
                     ws.Cells(Row, Column + 10).Value = percentChange
                     ws.Cells(Row, Column + 10).NumberFormat = "0.00%"
                 End If
               
                'Conditional Formatting - Highlight positive (Green) and negative (Red)
                If ws.Cells(Row, Column + 9).Value >= 0 Then
                     ' Set the Cell Color to Green
                     ws.Cells(Row, Column + 9).Interior.ColorIndex = 4
                 Else
                     ' Set the Cell Color to Red
                     ws.Cells(Row, Column + 9).Interior.ColorIndex = 3
                 End If
                 
                 'Add Total Volume
                 TotalTickerVolume = TotalTickerVolume + ws.Cells(i, Column + 6).Value
                 'MsgBox ("Total Volume: " & TotalTickerVolume)
                ws.Cells(Row, Column + 11).Value = TotalTickerVolume
                
                'Reset initial open price
                openPrice = ws.Cells(i + 1, Column + 2).Value
                
                 'Reset Total Stock Volume
                 TotalTickerVolume = 0
                
                'Add one row to summary table
                Row = Row + 1
            Else
                'Add the cells that have the same ticker
                TotalTickerVolume = TotalTickerVolume + ws.Cells(i, Column + 6).Value
            End If
                
        Next i
    
    
       '###### Challenge ######
          
        pcLastRow = ws.Cells(Rows.Count, Column + 11).End(xlUp).Row
        'MsgBox ("pcLastRow: " & pcLastRow)
        temp = Application.WorksheetFunction.Max(ws.Range("K2:K" & pcLastRow))
  
        'Retrieve stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
         For i = 2 To pcLastRow
            If ws.Range("K" & i).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & pcLastRow)) Then
                ws.Range("P2").Value = ws.Range("I" & i).Value
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("Q2").NumberFormat = "0.00%"
            ElseIf ws.Range("K" & i).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & pcLastRow)) Then
                ws.Range("P3").Value = ws.Range("I" & i).Value
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("Q3").NumberFormat = "0.00%"
            ElseIf ws.Range("L" & i).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & pcLastRow)) Then
                  ws.Range("P4").Value = ws.Range("I" & i).Value
                  ws.Range("Q4").Value = ws.Range("L" & i).Value
                
            End If
        Next i
          
       ' Format Table Columns To Auto Fit
         ws.Columns("I:Q").AutoFit
    Next ws
End Sub


