Sub GenerateSummary()

    Dim ticker As String
    Dim openingValue, closingValue As Double
    Dim stockVolume As Long
    Dim rowCounter, startRow As Integer
    Dim yearlyChangeRange As Range
    
    For Each ws In Worksheets
        Name = ws.Name
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        rowCounter = 2
        
        'Set the summary table header J to M columns
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percentage Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        
        For Row = 2 To lastrow
            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
                
                closingValue = ws.Cells(Row, 6).Value
                            
                'Set the values in the summary area
                ws.Cells(rowCounter, 10).Value = ws.Cells(Row, 1).Value
                ws.Cells(rowCounter, 11).Value = closingValue - openingValue
                
                If openingValue = 0 Then
                    ws.Cells(rowCounter, 12).Value = 0
                Else
                    ws.Cells(rowCounter, 12).Value = ws.Cells(rowCounter, 11).Value / openingValue
                    ws.Cells(rowCounter, 12).NumberFormat = "0.00%"
                End If
                            
                ws.Cells(rowCounter, 13).Value = "=SUM(" & Range(ws.Cells(startRow, 7), ws.Cells(Row, 7)).Address(False, False) & ")"
                            
                'Set/reset values for next entry
                rowCounter = rowCounter + 1
                openingValue = 0
                startRow = 0
            
            Else
                If openingValue = 0 Then
                    openingValue = ws.Cells(Row, 3).Value
                    startRow = Row
                End If
            End If
                
        Next Row
        
        'Conditional formatting for yearly change
        Set yearlyChangeRange = ws.Range("K2:K" & rowCounter)
        yearlyChangeRange.FormatConditions.Delete
        Set condition1 = yearlyChangeRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
        condition1.Interior.Color = vbRed
        Set condition2 = yearlyChangeRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        condition2.Interior.Color = vbGreen
       
        
        'Bonus Computation
        
        'Set the summary table header P,Q,R columns
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        
        ws.Range("R2").Value = "=MAX(L2:L" & rowCounter & ")"
        ws.Range("R3").Value = "=MIN(L2:L" & rowCounter & ")"
        ws.Range("R4").Value = "=MAX(M2:M" & rowCounter & ")"
     
        For Row = 2 To rowCounter
            If Cells(Row, 12).Value = ws.Range("R2").Value Then
                Range("Q2").Value = Cells(Row, 10).Value
            End If
            If Cells(Row, 12).Value = ws.Range("R3").Value Then
                Range("Q3").Value = Cells(Row, 10).Value
            End If
            If Cells(Row, 13).Value = ws.Range("R4").Value Then
                Range("Q4").Value = Cells(Row, 10).Value
            End If
        Next Row
        
        'Set style for the Greatest Values
        ws.Range("P1:R4").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ws.Range("P1:R4").Interior.ColorIndex = 34
        ws.Cells.EntireColumn.AutoFit
        
    Next ws

End Sub

