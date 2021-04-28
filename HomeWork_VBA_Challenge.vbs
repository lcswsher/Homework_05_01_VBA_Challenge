Sub VBAHomework()

'Loop through each worksheet in this workbook
For Each ws In Worksheets

'To create Column Headers for Stock Ticker, Yearly Price Change, percentage change, and Sum of Stock Volume
ws.Range("I1") = "Stock Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'To create Column Headers for Greatest increase, Greatest decrease, and Greatest Total volume
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'To Assign variables as either string long or Double
Dim StockTicker As String
Dim YearlyChange As Long
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim SummaryTableRow As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim ChangePrice As Double
Dim LastRow As Long

'To Assign variable for first opening ticker price
Dim OpeningPriceNumber As Long

'To Assign variables for Greatest Increase, Greatest decrease, and total volume
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

'Summary row starts at two and for each i loop will add a 1 count to the summarytablerow
SummaryTableRow = 2

'Summary row starts at two and for each i loop will add a 1 count to the OpenPriceNumber
OpeningPriceNumber = 2


'Calculates what the last row number in Column A is.
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'i variable represents Row number starts at 2 due to row 1 header on each sheet of workbook
    For i = 2 To LastRow

    'To summarize the volumes in column G for each individual stock symbol in Col A
    
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    'On change of Stock Ticker in Column A retrieve the ticker symbols in the sheet in column I
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
        'Set the StockTickerName
        StockTicker = ws.Cells(i, 1).Value
          
        'Testing - Message Box on change in Stock Ticker
        'MsgBox (Cells(i, 1).Value)
                       
        'Stock Ticker symbol to column I of each worksheet
        ws.Range("I" & SummaryTableRow).Value = StockTicker
        
        'Stock Volume is added from each cell while in loop - summarized by ticker symbol - to col L for each worksheet
        ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
        
        'To re-establish a zero balance for summary of stock volumes
        TotalStockVolume = 0
                
        'OpenPrice, ClosePrice, ChangePrice ("Ending Close Price" minus "Beginning Open Price" = "YearlyChange" in price)
        'OpenPrice
        OpenPrice = ws.Range("C" & OpeningPriceNumber).Value
        
        'ClosePrice
        ClosePrice = ws.Range("F" & i).Value
        
        'ChangePrice calculation and interior format of price
        ChangePrice = ClosePrice - OpenPrice
        ws.Range("J" & SummaryTableRow).Value = ChangePrice
        
            'Nested if statement for conditional formating for column K - percent change
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
               ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
            Else
               ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    
            End If
        
            'Percent Change calculation
            If OpenPrice = 0 Then
               PercentChange = 0
            Else
               OpenPrice = ws.Range("C" & OpeningPriceNumber)
               PercentChange = ChangePrice / OpenPrice
            
               ws.Range("K" & SummaryTableRow).Value = PercentChange
               ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

            End If
        
                       
        'increment 1 for SummaryTableRow and OpeningPriceNumber
        SummaryTableRow = SummaryTableRow + 1
        
        'Need to add the i count value to the opening price count to capture first price for stockticker
        OpeningPriceNumber = i + 1
                
    End If

    Next i

    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

    For i = 2 To LastRow
         If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            ws.Range("P2").Value = ws.Range("I" & i).Value
         End If
         
         If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
         End If
                              
         If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
         End If
                              
    Next i

'To properly fit data values into columns and to format number/percentages and center text

    ws.Columns("I:Q").AutoFit
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "0"
    ws.Range("Q1").HorizontalAlignment = xlCenter
            
Next ws
    
End Sub
