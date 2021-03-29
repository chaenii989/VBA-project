Sub StockMarket():

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    
    TotalVolume = 0
    
    For Each ws In Worksheets
    
        Dim Summary_Row As Integer
    
        Summary_Row = 2
    
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
    
        ws.Activate
    
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Chage"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        OpeningPrice = ws.Cells(2, "C").Value
            
            For i = 2 To lastrow
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    Ticker = ws.Cells(i, 1).Value
            
                    TotalVolume = TotalVolume + ws.Cells(i, "G").Value
                    
                    ClosingPrice = ws.Cells(i, "F").Value
    
                    YearlyChange = ClosingPrice - OpeningPrice
                            
                            If OpeningPrice <> 0 Then
                            
                            PercentChange = (ClosingPrice / OpeningPrice) - 1
                            
                            End If
                        
                    ws.Range("I" & Summary_Row).Value = Ticker
                    ws.Range("J" & Summary_Row).Value = YearlyChange
                        
                            If ws.Range("J" & Summary_Row).Value >= 0 Then
                            
                            ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                            
                            ElseIf ws.Range("J" & Summary_Row).Value < 0 Then
                            
                            ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                            
                            End If
                    
                    ws.Range("K" & Summary_Row).Value = PercentChange
                    ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                    ws.Range("L" & Summary_Row).Value = TotalVolume
                       
                    Summary_Row = Summary_Row + 1
                    TotalVolume = 0
                    OpeningPrice = ws.Cells(i + 1, "C").Value
            
                Else
                
                    TotalVolume = TotalVolume + ws.Cells(i, "G").Value
            
                End If
        
            Next i
       
            Dim GreatestIncrease As Double
            Dim GreatestDecrease As Double
            Dim GreatestTotal As Double
       
            GreatestIncrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
            ws.Range("P2") = ws.Cells(GreatestIncrease + 1, 9)
            ws.Cells(2, "Q").Value = WorksheetFunction.Max(ws.Range("K:K"))
            ws.Cells(2, "Q").NumberFormat = "0.00%"
       
            GreatestDecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
            ws.Range("P3") = ws.Cells(GreatestDecrease + 1, 9)
            ws.Cells(3, "Q").Value = WorksheetFunction.Min(ws.Range("K:K"))
            ws.Cells(3, "Q").NumberFormat = "0.00%"
       
            GreatestTotal = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
            ws.Range("P4") = ws.Cells(GreatestTotal + 1, 9)
            ws.Cells(4, "Q").Value = WorksheetFunction.Max(ws.Range("L:L"))
       
            ws.Cells.Columns.AutoFit
       
       Next ws
       
End Sub