# VBA-Challenge
Sub VBAChallenge()
    'Set Dimensions
    Dim Total As Double
    Dim RowIndex As Long
    Dim Change As Double
    Dim ColumnIndex As Integer
    Dim Start As Long
    Dim RowCount As Long
    Dim PercentChange As Double
    Dim Days As Integer
    Dim DailyChange As Single
    
    For Each ws In Worksheets
        ColumnIndex = 0
        Total = 0
        Change = 0
        Start = 2
        DailyChange = 0
        
      'Row Titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Value"
        
        RowCount = ws.UsedRange.Rows.Count
        
        For RowIndex = 2 To RowCount
            'Ticker changes 
            If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
                
                'Store Results in variable
                Total = Total + ws.Cells(RowIndex, 7).Value
                
                If Total = 0 Then
                    'Print the results
                    ws.Range("I" & 2 + ColumnIndex).Value = Cells(RowIndex, 1).Value
                    ws.Range("J" & 2 + ColumnIndex).Value = 0
                    ws.Range("K" & 2 + ColumnIndex).Value = "% & 0"
                    ws.Range("L" & 2 + ColumnIndex).Value = 0
                Else
                    If ws.Cells(Start, 3) = 0 Then
                        For Find_Value = Start To RowIndex
                            If ws.Cells(Find_Value, 3).Value <> 0 Then
                                Start = Find_Value
                                Exit For
                            End If
                        Next Find_Value
                    End If
                    
                    Change = (ws.Cells(RowIndex, 6) - ws.Cells(Start, 3))
                    PercentChange = Change / ws.Cells(Start, 3)
                    
                    Start = RowIndex + 1
                    
                    ws.Range("I" & 2 + ColumnIndex) = ws.Cells(RowIndex, 1).Value
                    ws.Range("J" & 2 + ColumnIndex) = Change
                    ws.Range("J" & 2 + ColumnIndex).NumberFormat = "0.00"
                    ws.Range("K" & 2 + ColumnIndex).Value = PercentChange
                    ws.Range("K" & 2 + ColumnIndex).NumberFormat = "0.00"
                    ws.Range("L" & 2 + ColumnIndex).Value = Total
                    
                    Select Case Change
                        Case Is > 0
                            ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 3
                        Case Else
                        ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 0
                    End Select
                        
                            
                End If
                
                Total = 0
                Change = 0
                ColumnIndex = ColumnIndex + 1
                Days = 0
                DailyChange = 0
                
            Else
            'Same Ticker
                Total = Total + ws.Cells(RowIndex, 7).Value
            
            End If
            
        
        
        Next RowIndex
        
        'take the max and min and place them in a separate part of worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
        
        Increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        Decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
        
        ws.Range("P2") = ws.Cells(Increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(Decrease_Number + 1, 9)
        ws.Range("P3") = ws.Cells(volume_number + 1, 9)
         

    
    Next ws
End Sub
