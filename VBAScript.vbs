Attribute VB_Name = "Module1"
Sub ProcessStockData()
    Dim tickerName As String
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim summaryTickerRow As Integer
    
    totalVolume = 0
    summaryTickerRow = 2
    openPrice = Cells(2, 3).Value

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
    If IsEmpty(Cells(i + 1, 1).Value) Or Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        tickerName = Cells(i, 1).Value
        closePrice = Cells(i, 6).Value
        yearlyChange = closePrice - openPrice
        percentChange = IIf(openPrice = 0, 0, yearlyChange / openPrice)
        
        totalVolume = totalVolume + Cells(i, 7).Value

        With Range("I" & summaryTickerRow)
            .Value = tickerName
            .Offset(0, 3).Value = totalVolume
            .Offset(0, 1).Value = yearlyChange
            .Offset(0, 2).Value = percentChange
            .Offset(0, 2).NumberFormat = "0.00%"
        End With

        summaryTickerRow = summaryTickerRow + 1
        totalVolume = 0
        openPrice = Cells(i + 1, 3)
    Else
        totalVolume = totalVolume + Cells(i, 7).Value
    End If
    
Next i

  lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    
    
End Sub



