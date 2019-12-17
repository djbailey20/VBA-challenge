Attribute VB_Name = "Module1"
Sub stock()
    Dim Worksheet
    Dim RowNumber As Integer
    Dim total As Variant
    Dim YearStart As Double
    Dim YearEnd As Double
    Dim HighestIncrease As Double
    Dim HITicker As String
    Dim HighestDecrease As Double
    Dim HDTicker As String
    Dim HighestVolume As Double
    Dim HVTicker As String
    HighestIncrease = 0
    HighestDecrease = 0
    HighestVolume = 0
    
    For Each WS In Worksheets
        WS.Range("J1").Value = "Ticker"
        WS.Range("K1").Value = "Yearly Change"
        WS.Range("L1").Value = "Percent Change"
        WS.Range("M1").Value = "Total Stock Volume"
        RowNumber = 2
        total = 0
        For i = 2 To WS.Cells(Rows.Count, 1).End(xlUp).Row
            If WS.Cells(i - 1, 1).Value <> WS.Cells(i, 1).Value Then
                YearStart = WS.Cells(i, 3).Value
            End If
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                WS.Range("J" & RowNumber).Value = WS.Cells(i, 1).Value
                total = total + WS.Cells(i, 7).Value
                WS.Range("K" & RowNumber).Value = WS.Cells(i, 6).Value - YearStart
                If WS.Range("K" & RowNumber).Value < 0 Then
                    WS.Range("K" & RowNumber).Interior.Color = RGB(255, 0, 0)
                ElseIf WS.Range("K" & RowNumber).Value > 0 Then
                    WS.Range("K" & RowNumber).Interior.Color = RGB(0, 255, 0)
                End If
                If YearStart <> 0 Then
                    WS.Range("L" & RowNumber).Value = WS.Range("K" & RowNumber).Value / YearStart
                Else
                    WS.Range("L" & RowNumber).Value = 0
                End If
                WS.Range("L" & RowNumber).NumberFormat = "0.00%"
                WS.Range("M" & RowNumber).Value = total
                total = 0
                RowNumber = RowNumber + 1
                
            Else
                total = total + WS.Cells(i, 7).Value
            End If
        Next i
        
        For i = 2 To WS.Cells(Rows.Count, 10).End(xlUp).Row
            
            If WS.Cells(i, 12).Value > HighestIncrease Then
                HighestIncrease = WS.Cells(i, 12).Value
                HITicker = WS.Cells(i, 10).Value
            End If
            If WS.Cells(i, 12).Value < HighestDecrease Then
                HighestDecrease = WS.Cells(i, 12).Value
                HDTicker = WS.Cells(i, 10).Value
            End If
            If WS.Cells(i, 13).Value > HighestVolume Then
                HighestVolume = WS.Cells(i, 13).Value
                HVTicker = WS.Cells(i, 10).Value
            End If
        Next i
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("O3").Value = "Greatest % Decrease"
        WS.Range("O4").Value = "Greatest Total Volume"
        
        WS.Range("P1").Value = "Ticker"
        WS.Range("P2").Value = HITicker
        WS.Range("P3").Value = HDTicker
        WS.Range("P4").Value = HVTicker
        
        WS.Range("Q1").Value = "Value"
        WS.Range("Q2").Value = HighestIncrease
        WS.Range("Q3").Value = HighestDecrease
        WS.Range("Q4").Value = HighestVolume
        
        WS.Range("Q2").NumberFormat = "0.00%"
        WS.Range("Q3").NumberFormat = "0.00%"
        HighestIncrease = 0
        HighestDecrease = 0
        HighestVolume = 0
    Next WS
End Sub
