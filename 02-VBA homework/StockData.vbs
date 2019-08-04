Attribute VB_Name = "Module1"
Sub StockData()

'Dim up some Vars Brah!
Dim Ticker, NextTicker As String
Dim Vol, TotalVol, SumRow, LastRow, openval, high, low, closeval, Change As Double
Dim ws As Worksheet

'Loop through each worksheet and do some stuff

For Each ws In Worksheets

'Get Last Row & Last Column of the Current Sheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Initialize the Totals Bucket & Set the Summary Row Start
     TotalVol = 0
     SumRow = 2
     
'Create and format the header row of the summary table
     ws.Range("M2") = Range("c2")
     ws.Range("K1") = "Ticker"
     ws.Range("L1") = "Total Volume"
     ws.Range("M1") = "Open Beginning of the Year"
     ws.Range("N1") = "Close End of the Year"
     ws.Range("O1") = "Change"
     ws.Range("P1") = "Change %"
     
          ws.Range("K1:P1").Font.FontStyle = "Bold"
          ws.Range("K1:P1").Interior.Color = RGB(200, 200, 200)
          ws.Columns("K:P").AutoFit
          ws.Range("M:O").Style = "Currency"
          ws.Range("P:P").Style = "Percent"

'Loop Through Tickers for each change

     For i = 2 To LastRow

'Define the data to keep it clean
     Ticker = ws.Cells(i, 1)
     NextTicker = ws.Cells(i + 1, 1)
     openval = ws.Cells(i + 1, 3)
     closeval = ws.Cells(i, 6)
     Vol = ws.Cells(i, 7)

'Looking for the change in Ticker to compare values
          If Ticker <> NextTicker Then
               'ws.Range("J" & SumRow) = ws.Name
               ws.Range("K" & SumRow) = Ticker
               ws.Range("L" & SumRow) = TotalVol + Vol
               ws.Range("M" & SumRow + 1) = openval
               ws.Range("N" & SumRow) = closeval
               ws.Range("O" & SumRow) = ws.Range("N" & SumRow) - ws.Range("M" & SumRow)
                    
                    'Checking for a null or zero before dividing
                    If ws.Range("M" & SumRow) = "" Or ws.Range("M" & SumRow) = 0 Then
                         ws.Range("P" & SumRow) = 0
                    Else
                         ws.Range("P" & SumRow) = ws.Range("O" & SumRow) / ws.Range("M" & SumRow)
                    End If
               
                    'Format Cell based on value in Change%
                    If ws.Range("P" & SumRow) > 0 Then
                         ws.Range("P" & SumRow).Interior.ColorIndex = 4
                    Else
                         ws.Range("P" & SumRow).Interior.ColorIndex = 3
                    End If
               
               SumRow = SumRow + 1
               TotalVol = 0
               
               Else
               TotalVol = TotalVol + Vol
               
               
          End If

     Next i

Next ws


End Sub

