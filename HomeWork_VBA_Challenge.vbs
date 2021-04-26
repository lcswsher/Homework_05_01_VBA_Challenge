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

'Calculate what the last row number in Column A

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'i represents Row number

    For i = 2 To LastRow

    'Add Stock Ticker to

    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

    'On change of Stock Ticker in Column A retrieve the ticker symbol

    Next i
    
Next ws
    
    
End Sub
