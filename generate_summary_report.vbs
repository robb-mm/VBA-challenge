Option Explicit

Sub generate_summary_report()

Dim i As Long, j As Integer: j = 1
Dim OpenPrice() As Double, ClosePrice() As Double
Dim QuarterlyChange As Double, PercentChange As Double
Dim LastRow As Long, Vol As LongLong
Dim GreatestPriceIncrease As Double, GreatestPriceDecrease As Double, GreatestTotalVol As LongLong
Dim TopStock As String, BottomStock As String, TopStockByVol As String
Dim StockSymbol() As String, Volume() As Double
Dim ws As Worksheet

GreatestPriceIncrease = 0
GreatestPriceDecrease = 0
GreatestTotalVol = 0

For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
Vol = 0

For i = 2 To LastRow + 1

    ' Open Price
    If Vol = 0 Then
        ReDim Preserve StockSymbol(1 To j)
        ReDim Preserve OpenPrice(1 To j)
        
        StockSymbol(j) = ws.Cells(i, 1).Value
        OpenPrice(j) = ws.Cells(i, 3).Value
    End If

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ReDim Preserve ClosePrice(1 To j)
        ReDim Preserve Volume(1 To j)
    
        'Close Price
        ClosePrice(j) = ws.Cells(i, 6).Value
        
        'Add to the Volume total
        Volume(j) = Vol + ws.Cells(i, 7).Value
        
        'Reset the Volume total
        Vol = 0
        j = j + 1
    Else
        'Add to the Volume total
        Vol = Vol + ws.Cells(i, 7).Value
    End If
    
Next i

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Quarterly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

For i = 1 To j - 1
    QuarterlyChange = CCur(ClosePrice(i)) - CCur(OpenPrice(i))
    PercentChange = QuarterlyChange / CCur(OpenPrice(i))
    
    ws.Range("I" & (i + 1)) = StockSymbol(i)
    ws.Range("J" & (i + 1)) = QuarterlyChange
    ws.Range("K" & (i + 1)) = PercentChange
    ws.Range("L" & (i + 1)) = Volume(i)
    
    ws.Range("K" & (i + 1)).NumberFormat = "0.00%"
    If (QuarterlyChange > 0) Then
        ws.Range("J" & (i + 1)).Interior.ColorIndex = 4
        ws.Range("K" & (i + 1)).Interior.ColorIndex = 4
    ElseIf (QuarterlyChange < 0) Then
        ws.Range("J" & (i + 1)).Interior.ColorIndex = 3
        ws.Range("K" & (i + 1)).Interior.ColorIndex = 3
    End If
    
    
    If (GreatestPriceIncrease < PercentChange) Then
        GreatestPriceIncrease = PercentChange
        TopStock = StockSymbol(i)
    End If
    
    If (GreatestPriceDecrease > PercentChange) Then
        GreatestPriceDecrease = PercentChange
        BottomStock = StockSymbol(i)
    End If
    
    If (GreatestTotalVol < Volume(i)) Then
        GreatestTotalVol = Volume(i)
        TopStockByVol = StockSymbol(i)
    End If
    
Next i

' show the greatest variables
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P2") = TopStock
ws.Range("P3") = BottomStock
ws.Range("P4") = TopStockByVol
ws.Range("Q2") = GreatestPriceIncrease
ws.Range("Q3") = GreatestPriceDecrease
ws.Range("Q4") = GreatestTotalVol

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

' reset the greatest variables
GreatestPriceIncrease = 0
GreatestPriceDecrease = 0
GreatestTotalVol = 0

ReDim StockSymbol(1 To 1)
ReDim OpenPrice(1 To 1)
ReDim ClosePrice(1 To 1)
ReDim Volume(1 To 1)
j = 1

Next ws

End Sub



