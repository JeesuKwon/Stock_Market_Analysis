Attribute VB_Name = "Module1"
Sub stockdata()
    'Loop for worksheets'
    For Each ws In Worksheets
        
        ''code for 1st table''
        
        'Declarations'
        Dim i, j As Long
        Dim tickcount As Long
        Dim lastrow As Long
        
        
        'Header'
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
        'set up the basic counts for loops'
        tickcount = 2
        j = 2
        
        'set up last row'
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'for loop start!'
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(tickcount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(tickcount, 10).Value = ws.Cells(i, 6) - Cells(j, 3).Value
                    If ws.Cells(tickcount, 10).Value > 0 Then
                        ws.Cells(tickcount, 10).Interior.ColorIndex = 4
                    Else:
                        ws.Cells(tickcount, 10).Interior.ColorIndex = 3
                    End If
                ws.Cells(tickcount, 11).Value = Format((ws.Cells(i, 6) - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value, "percent")
                ws.Cells(tickcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                tickcount = tickcount + 1
                j = i + 1
                
                
    
            
            End If
            
        Next i
        
        
        ''code for 2nd table''
        
        'new last row for second table'
        newll = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
        'Declarations'
        greatvol = ws.Cells(2, 12).Value
        greatinc = ws.Cells(2, 11).Value
        greatdec = ws.Cells(2, 11).Value
        
            'For loop start!'
            For i = 2 To newll
            
                If ws.Cells(i, 12).Value > greatvol Then
                greatvol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                
                Else
                greatvol = greatvol
                
                End If
                
                
            
                If ws.Cells(i, 11).Value > greatinc Then
                greatinc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                
                Else
                greatinc = greatinc
                
                End If
                
                
                If ws.Cells(i, 11).Value < greatdec Then
                greatdec = Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                
                Else
                greatdec = greatdec
                
                End If
            
            
            Next i
            
        'print out value on the table'
        ws.Cells(2, 17).Value = Format(greatinc, "percent")
        ws.Cells(3, 17).Value = Format(greatdec, "percent")
        ws.Cells(4, 17).Value = Format(greatvol, "scientific")
        
    Next ws
    

End Sub
