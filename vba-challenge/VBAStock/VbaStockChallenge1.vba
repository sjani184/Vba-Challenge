Attribute VB_Name = "Module1"
Sub StockTracker()
Dim Ticker As String
Dim Summary_Stock_Row As Long
Dim LastRow As Long
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As String
Dim StockClose As Double
Dim StockOpen As Double

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"

    Summary_Stock_Row = 2

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    YearlyChange = 0

    TotalVolume = 0

    StockClose = 0


    For i = 2 To LastRow
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

            StockOpen = ws.Cells(i, 3).Value

        End If

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            StockClose = ws.Cells(i, 6).Value
            
            YearlyChange = StockClose - StockOpen
            
            If StockOpen = 0 Then
            
                PercentChange = 0
            Else
            
                PercentChange = (StockClose - StockOpen) / StockOpen
            
            End If

            ws.Range("J" & Summary_Stock_Row).Value = YearlyChange

            ws.Range("K" & Summary_Stock_Row).Value = FormatPercent(PercentChange, 1)

            Summary_Stock_Row = Summary_Stock_Row + 1


        End If
        
            If ws.Cells(i, 10).Value > 0 Then
            
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
            
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
    Next i
    
    
    Summary_Stock_Row = 2


    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value

            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            ws.Range("I" & Summary_Stock_Row).Value = Ticker

            ws.Range("L" & Summary_Stock_Row).Value = TotalVolume

            Summary_Stock_Row = Summary_Stock_Row + 1

            TotalVolume = 0

        Else

           TotalVolume = TotalVolume + ws.Cells(i, 7).Value


        End If


     Next i

Next ws

End Sub
