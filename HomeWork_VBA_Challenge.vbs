Sub VBAHomework()

'Loop through each worksheet in this workbook
For Each ws In Worksheets

'To create Column Headers for Stock Ticker, Yearly Price Change, percentage change, and Sum of Stock Volume

ws.Range("I1") = "Stock Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'To Assign variables as either string or long

Dim StockTicker As String
Dim YearlyChange As Long
Dim PercentChange As Long
Dim TotalStockVolume As Long
Dim SummaryTableRow As Long

'Summary row starts at two and for each i loop will add a 1 count to the summarytablerow
SummaryTableRow = 2

'Calculates what the last row number in Column A

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'i variable represents Row number

    For i = 2 To LastRow

    'To summarize the volumes in column G for each individual stock symbol in Col A
    
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    'On change of Stock Ticker in Column A retrieve the ticker symbols in the sheet in column I
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               
        'Set the StockTickerName
        StockTicker = ws.Cells(i, 1).Value
          
        'Testing - Message Box on change in Stock Ticker
        'MsgBox (Cells(i, 1).Value)
                       
        'Stock Ticker symbol to column I of each worksheet
        ws.Range("I" & SummaryTableRow).Value = StockTicker
        
        'Stock Volume is added from each cell while in loop summarized by ticker symbol - to col L of each worksheet
        ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
        
        SummaryTableRow = SummaryTableRow + 1
                                   
                       
    End If

    Next i
    
Next ws
    
End Sub
