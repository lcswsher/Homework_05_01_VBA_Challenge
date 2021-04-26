Sub VBAHomework()

'Loop through each worksheet in this workbook
    For Each ws In Worksheets

'To create Column Headers for Stock Ticker, Yearly Price Change, percentage change, and Sum of Stock Volume

ws.Range("I1") = "Stock Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

Dim StockTicker As String
Dim YearlyChange As Long
Dim PercentChange As Long
Dim TotalStockVolume As Long

    Next ws
    
End Sub
