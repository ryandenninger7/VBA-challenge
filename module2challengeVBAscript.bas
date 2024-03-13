Attribute VB_Name = "Module1"
Sub CalculateAllStockData()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        CalculateStockData ws
    Next ws
End Sub

Sub CalculateStockData(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim startPrice As Double
    Dim endPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim ticker As String
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    startPrice = ws.Cells(2, 3).Value
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            endPrice = ws.Cells(i, 6).Value
            
            yearlyChange = endPrice - startPrice
            If startPrice <> 0 Then
                percentChange = (yearlyChange / startPrice) * 100
            Else
                percentChange = 0
            End If
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                tickerIncrease = ticker
            ElseIf percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                tickerDecrease = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                tickerVolume = ticker
            End If
            
            Dim outputRow As Long
            outputRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row + 1
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = yearlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = totalVolume
            
            totalVolume = 0
            If i + 1 <= lastRow Then
                startPrice = ws.Cells(i + 1, 3).Value
            End If
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ws.Cells(2, 15).Value = tickerIncrease
    ws.Cells(2, 16).Value = greatestIncrease
    ws.Cells(3, 15).Value = tickerDecrease
    ws.Cells(3, 16).Value = greatestDecrease
    ws.Cells(4, 15).Value = tickerVolume
    ws.Cells(4, 16).Value = greatestVolume
End Sub


