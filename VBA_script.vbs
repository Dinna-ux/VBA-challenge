Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data():

    For Each ws In Worksheets
    
    'variables
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim T_Count As Long
        Dim L_RowA As Long
        Dim L_RowI As Long
        Dim P_Change As Double
        Dim G_Incr As Double
        Dim G_Decr As Double
        Dim G_Vol As Double
        
        'worksheet name
        WorksheetName = ws.Name
        
        'Columns headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Rows
        T_Count = 2
        
        'start row
        j = 2
        
        'Last cell in the first column
        L_RowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Initialise the loop to scan through rows
            For i = 2 To L_RowA
            
            'Compare the cells values for ticker and add ticker value to column i
            
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ws.Cells(T_Count, 9).Value = ws.Cells(i, 1).Value
                
                
                'Find the Yearly change and populate data in J.
                
                ws.Cells(T_Count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                
              'Applying conditional formating by setting the bg color to red(3),and green(4).
             If ws.Cells(T_Count, 10).Value < 0 Then
                
                'Red background
                    ws.Cells(T_Count, 10).Interior.ColorIndex = 3
                
                     Else
                
                       'Green background
                       ws.Cells(T_Count, 10).Interior.ColorIndex = 4
                
                    End If
                    
              'Percentage change column calculation and formating
              If ws.Cells(j, 3).Value <> 0 Then
                P_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                 ws.Cells(T_Count, 11).Value = Format(P_Change, "Percent")
                    
                    Else
                    
                      ws.Cells(T_Count, 11).Value = Format(0, "Percent")
                    
                    End If
                     'Total volume
                ws.Cells(T_Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                    'Creating a new start row for ticker
                     T_Count = T_Count + 1
                
                      j = i + 1
                
                End If
            
            Next i
            
        'Checking for the last populated cell
        L_RowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        G_Vol = ws.Cells(2, 12).Value
        G_Incr = ws.Cells(2, 11).Value
        G_Decr = ws.Cells(2, 11).Value
        
            For i = 2 To L_RowI
            
            If ws.Cells(i, 12).Value > G_Vol Then
                G_Vol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                  Else
                
                     G_Vol = G_Vol
                
             End If
                
               'Compare values for greatest increase with the next value
             If ws.Cells(i, 11).Value > G_Incr Then
                G_Incr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                  Else
                
                    G_Incr = G_Incr
                
                End If
                
                'Compare values for greatest decrease with the next value
            If ws.Cells(i, 11).Value < G_Decr Then
                G_Decr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                    Else
                
                        G_Decr = G_Decr
                
                End If
                
             'Populating summary results in the worksheet
            ws.Cells(2, 17).Value = Format(G_Incr, "Percent")
            ws.Cells(3, 17).Value = Format(G_Decr, "Percent")
            ws.Cells(4, 17).Value = Format(G_Vol, "Scientific")
            
            Next i
            
             Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub

