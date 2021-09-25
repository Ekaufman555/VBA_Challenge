Attribute VB_Name = "Module1"
Sub StockMarket1():

    For Each ws In Worksheets
    
        ' First we need to declare variables
        Dim StockTick As String 'Stock Ticker
        Dim StockOpen As Double 'Opening Price
        Dim StockClose As Double 'Closing Price
        Dim Volume As LongLong 'Individual volume may not be necessary
        Dim TotVolume As LongLong 'Total volume across ticker
        Dim SummaryTable As Integer
        Dim start As Long
        start = 2
        
        Dim Lastrow As Long
        
        ws.Cells(1, 10).Value = "Tickers"
        ws.Cells(1, 11).Value = "Change in Value"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Volume"
        SummaryTable = 2
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For Row = 2 To Lastrow
            If ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1).Value Then
                'Set the Opening value for the ticker, since Row 1 are headers this will make each subsequent ticker change modify the opening after the change
                If ws.Range("C" & Row).Value = 0 Then
                    For x = start To Row
                    
                        If ws.Range("C" & x).Value <> 0 Then
                            StockOpen = ws.Range("C" & x).Value
                        End If
                        Exit For
                        
                    Next x
                    
                Else
                    StockOpen = ws.Range("C" & Row).Value
                
                    
                End If
                
                
                Volume = Volume + ws.Range("G" & Row).Value
                
                
            ElseIf ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
                
               'Reset Stock Ticker
                StockTick = ws.Range("A" & Row).Value
                Volume = Volume + ws.Range("G" & Row).Value
                StockClose = ws.Range("F" & Row).Value
                
               'J Column will start stock tickers
                ws.Range("J" & SummaryTable).Value = StockTick
                ws.Range("K" & SummaryTable).Value = StockOpen - StockClose
                If ws.Range("K" & SummaryTable).Value > 0 Then
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 4
                ElseIf ws.Range("K" & SummaryTable).Value < 0 Then
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 3
                Else
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 0
                
                
                
                
                End If
                
                ws.Range("L" & SummaryTable).Value = Range("K" & SummaryTable).Value / StockOpen
                
                ws.Range("L" & SummaryTable).Style = "Percent"
                ws.Range("M" & SummaryTable).Value = Volume
                SummaryTable = SummaryTable + 1
                Volume = 0
                
                
            Else
                 Volume = Volume + ws.Range("G" & Row).Value
        
            End If
            
        Next Row
        
    Next ws
    
End Sub
