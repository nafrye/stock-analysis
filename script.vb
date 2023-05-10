Sub Ticker()
For Each ws In Worksheets
    Dim Worksheet_name As String
    Worksheet_name = ws.Name

    'define LastRow
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (LastRow)


    'write label for column I
    ws.Cells(1, 9).Value = "Ticker"
    
    'write label for column J
    ws.Cells(1, 10).Value = "Yearly Change"
    
    'write label for column K
    ws.Cells(1, 11).Value = "Percent Change"
    
    'write label for column L
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'strategy: read the ticker from first column, see if it is equal to the ticker in the row below, store opening and closing vlaue as variable
    
    
    'define number to write ticker in next blank row in ticker column
    Dim j As Integer
    
    'define variable for opening value on first of year
    Dim o As Double
    
    'define variable for closing value on last day of year
    Dim c As Double
    
    'define variable for change in price
    Dim change As Double
    
    'define variable for percent change in price
    Dim per_change As Double
    
    'define variable for total volume
    Dim tot_vol As Double

                
        
    j = 2
    For i = 2 To LastRow
        'Range("G" & i) = CLng(Range("B" & i))
        
        If (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
            'get opening price
            o = ws.Cells(i, 3).Value
            
            'start volume again with initial value for new stock
            tot_vol = ws.Range("G" & i).Value
            'MsgBox (tot_vol)
        
      
        ElseIf (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
            tot_vol = tot_vol + ws.Cells(i, 7).Value

        ElseIf (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            'write ticker name
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
            
            'get closing price
            c = ws.Cells(i, 6).Value
            'compute change in price
            change = c - o
            
            'compute percent change in price
            per_change = change / o
            
            'write change in price
            ws.Cells(j, 10).Value = change
            
            'write percent change in price
            ws.Cells(j, 11).Value = per_change
            
            'add to total value
            tot_vol = tot_vol + ws.Range("G" & i).Value
            
            'write total volume
            ws.Cells(j, 12).Value = tot_vol
        
           
            j = j + 1
               
        End If
                            
        
    Next i
               
    'get number of j rows
    LastRowOut = ws.Cells(Rows.Count, "I").End(xlUp).Row
    'MsgBox (LastRowOut)
   
   
   'color cell based on whether positive or negative change
    For b = 2 To LastRowOut
    If (ws.Cells(b, 10).Value > 0) Then
        ws.Cells(b, 10).Interior.ColorIndex = 4
        
    ElseIf (ws.Cells(b, 10).Value < 0) Then
        ws.Cells(b, 10).Interior.ColorIndex = 3
    End If
    
    Next b
   
   
   'write percent change as percent
   'note: will need to change to .Range when scroll sheet by sheet
   ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
        
        
    'write greatest labels
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    
    'select maximums and write to cells
    ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRowOut))
    ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRowOut))
    ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRowOut))

    'change max and min percent change to percent
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'get ticker names for largest increase in percentage, largest decrease in percentage, and largest volume
    For p = 2 To LastRowOut
    If (ws.Cells(p, 11).Value = ws.Cells(2, 17).Value) Then
        ws.Cells(2, 16).Value = ws.Cells(p, 9).Value
        End If
    If (ws.Cells(p, 11).Value = ws.Cells(3, 17).Value) Then
        ws.Cells(3, 16).Value = ws.Cells(p, 9).Value
        End If
    If (ws.Cells(p, 12).Value = ws.Cells(4, 17).Value) Then
        ws.Cells(4, 16).Value = ws.Cells(p, 9).Value
        End If
    
    Next p
    
    
    Next ws
    


End Sub