Attribute VB_Name = "Module1"
Sub MYSDLoop():
    
    'Declaring variables for each ws
        Dim ws As Worksheet
        Dim i, j As Long
        Dim TickerSymbol, TickerOrg, TickerNew As Long
        Dim Percent As Double
        Dim GreaestInc, GreatestDec, GreatestVol As Double
    
    'Creating labels for each ws
        For Each ws In Worksheets
        
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
                    
        'Find last row of original Ticker column
            TickerOrg = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Sets the count of ticker to begin at row 2
            TickerSymbol = 2
             j = 2

        'Script loop for Ticker, Yearly Change, Percent Change and Total Stock Volume
        'Begin with finidng new Ticker column outcome, grouping like symbols
            For i = 2 To TickerOrg
        
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerSymbol, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate Yearly Change values based on the Ticker outcome based on sum of columns close & open
                ws.Cells(TickerSymbol, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
                        'Create Conditional Formatting based on the outcome of the yearly change, green (value > 0), red (value < 0)
                        If ws.Cells(TickerSymbol, 10).Value > 0 Then
                        ws.Cells(TickerSymbol, 10).Interior.ColorIndex = 4   
                Else
                ws.Cells(TickerSymbol, 10).Interior.ColorIndex = 3
                End If
                    
                'Caluclating Percent Change Column pulling data from percentage columns
                If ws.Cells(i, 3).Value <> 0 Then
                Percent = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                ws.Cells(TickerSymbol, 11).Value = Format(Percent, "Percent")
            
                Else
                ws.Cells(TickerSymbol, 11).Value = Format(0, "Percent")
                End If
        
                'Calculating Total Stock Volume column, pulling data from the original volume column
                ws.Cells(TickerSymbol, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

                'Start new row for every new Ticker
                TickerSymbol = TickerSymbol + 1
                j = i + 1
    
            End If
        
    Next i
        
    'Loop for summary table
        'Define which columns the summarized ticker data is being pulled from
        TickerNew = ws.Cells(Rows.Count, 9).End(xlUp).Row
        GreatestInc = ws.Cells(2, 11).Value
        GreatestDec = ws.Cells(2, 11).Value
        GreatestVol = ws.Cells(2, 12).Value
        
        'Pull from row 2 to bottom of the new ticker column
            For i = 2 To TickerNew
                'Calculated value for greatest total volume
                'Refer to Total Stock Volume column, place greatest value in "P4"
                    If ws.Cells(i, 12).Value > GreatestVol Then
                    GreatestVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                    Else
                    GreatestVol = GreatestVol
                    End If
                
                'Calculated value for greatest % increase
                'Refer to Percent Change column, place greatest % increase value in "P2"
                    If ws.Cells(i, 11).Value > GreatestInc Then
                    GreatestInc = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    GreatestInc = GreatestInc
                    End If
                
                'Calculated value for greatest % decrease
                'Refer to Percent Change column, place greatest % decrease value in "P3"
                    If ws.Cells(i, 11).Value < GreatestDec Then
                    GreatestDec = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    GreatestDec = GreatestDec
                    End If
                    
                'Formatting the outcome of the Ticker values from P2, P3, P4
                ws.Range("Q2").Value = Format(GreatestInc, "Percent")
                ws.Range("Q3").Value = Format(GreatestDec, "Percent")
                ws.Range("Q4").Value = Format(GreatestVol, "Scientific")
                
                Next i
                
            'Autofit all cells in each ws
            ws.Cells.EntireColumn.AutoFit
            ws.Cells.EntireRow.AutoFit
            
    Next ws
    
End Sub
