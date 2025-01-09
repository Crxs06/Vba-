Sub dateopen()
 
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim Volume As Double
    Dim QuarterlyChange As Double
    Dim Percentage As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolumeTicker As String
    Dim LastRow As Long
    Dim Outpow As Long
    Dim i As Long
    
    '   Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    ws.Activate
    
    ' Intialize variables
    Volume = 0
    OpeningPrice = ws.Cells(2, 3).Value
    OutputRow = 2
    
    ' Determinr last Row
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    '  Set headers for output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        GreatestIncrease = 0
        GraetestDecrease = 0
        GreatestVolume = 0
        
            '   Loop through each row of data
            For i = 2 To LastRow
            ' Check if ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosingPrice = ws.Cells(i, 6).Value
                Volume = Volume + ws.Cells(i, 7).Value
                
                ' Calculate quarterly change and percentage change
                QuarterlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentageChange = QuarterlyChange / OpeningPrice
                Else
                    PercentageChange = 0
                End If
                
                ' Output Results
                ws.Cells(OutputRow, 9).Value = Ticker
                ws.Cells(OutputRow, 10).Value = QuarterlyChange
                ws.Cells(OutputRow, 11).Value = PercentageChange
                ws.Cells(OutputRow, 12).Value = Volume
                
                ' Check for greatest increase, decrease,  and volume
                If PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    GreatestIncreaseTicker = Ticker
                End If
                    
                If PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    GreatestDecreaseTicker = Ticker
                End If
                
                If Volume > GreatestVolume Then
                    GreatestVolume = Volume
                    GreatestVolumeTicker = Ticker
                    
                    End If
                
             ' Reset for next ticker
             Volume = 0
             If i + 1 <= LastRow Then
                OpeningPrice = ws.Cells(i + 1, 3).Value
            End If
            OutputRow = OutputRow + 1
        Else
            ' Accumulate volume
            Volume = Volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ' Output summary table
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = GreatestIncreaseTicker
    ws.Cells(2, 16).Value = GreatestIncrease
    
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = GreatestDecreaseTicker
    ws.Cells(3, 16).Value = GreatestDecrease
    
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Value = GreatestVolumeTicker
    ws.Cells(4, 16).Value = GreatestVolume
Next ws
    End Sub
