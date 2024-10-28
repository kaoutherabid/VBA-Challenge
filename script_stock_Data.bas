Attribute VB_Name = "Module1"
Sub StockData()
    Dim ws As Worksheet
    Dim lastRow As Long, row As Long, TickerRow As Long
    Dim TickerName As String, StockVolumeTotal As Double
    Dim openprice As Double, closeprice As Double, qchange As Double
    Dim maxpchange As Double, minpchange As Double, maxvol As Double
    Dim maxpchangeindex As Long, minpchangeindex As Long, maxvolindex As Long
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        TickerRow = 2  ' Start at row 2 for new ticker data
        StockVolumeTotal = 0  ' Reset stock volume
        openprice = ws.Cells(2, 3).Value  ' Initialize open price
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row  ' Find the last row with data
        ' Populate headers in the current worksheet
        ' Add headers to the summary columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ' Process each row in the worksheet
        For row = 2 To lastRow
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                ' New Ticker found; store data in summary columns
                TickerName = ws.Cells(row, 1).Value
                ws.Cells(TickerRow, 9).Value = TickerName  ' Column I
                closeprice = ws.Cells(row, 6).Value  ' Closing price
                qchange = closeprice - openprice  ' Quarterly change
                ws.Cells(TickerRow, 10).Value = qchange  ' Column J
                ' Color code quarterly change
                If qchange > 0 Then
                    ws.Cells(TickerRow, 10).Interior.ColorIndex = 4  ' Green
                Else
                    ws.Cells(TickerRow, 10).Interior.ColorIndex = 3  ' Red
                End If
                ' Calculate and format percent change
                If openprice <> 0 Then
                    ws.Cells(TickerRow, 11).Value = qchange / openprice  ' Column K
                Else
                    ws.Cells(TickerRow, 11).Value = 0
                End If
                ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
                ' Add total stock volume
                StockVolumeTotal = StockVolumeTotal + ws.Cells(row, 7).Value
                ws.Cells(TickerRow, 12).Value = StockVolumeTotal  ' Column L
                ' Reset variables for next ticker
                StockVolumeTotal = 0
                TickerRow = TickerRow + 1  ' Move to the next row for ticker
                openprice = ws.Cells(row + 1, 3).Value  ' New open price
            Else
                ' Accumulate stock volume if ticker hasn't changed
                StockVolumeTotal = StockVolumeTotal + ws.Cells(row, 7).Value
            End If
        Next row
        ' Generate summary for the current worksheet
        maxpchange = WorksheetFunction.Max(ws.Range("K2:K" & ws.Cells(ws.Rows.Count, 11).End(xlUp).row))
        maxpchangeindex = WorksheetFunction.Match(maxpchange, ws.Range("K2:K" & ws.Cells(ws.Rows.Count, 11).End(xlUp).row), 0) + 1
        minpchange = WorksheetFunction.Min(ws.Range("K2:K" & ws.Cells(ws.Rows.Count, 11).End(xlUp).row))
        minpchangeindex = WorksheetFunction.Match(minpchange, ws.Range("K2:K" & ws.Cells(ws.Rows.Count, 11).End(xlUp).row), 0) + 1
        maxvol = WorksheetFunction.Max(ws.Range("L2:L" & ws.Cells(ws.Rows.Count, 12).End(xlUp).row))
        maxvolindex = WorksheetFunction.Match(maxvol, ws.Range("L2:L" & ws.Cells(ws.Rows.Count, 12).End(xlUp).row), 0) + 1
        ' Output the summary
        ws.Range("O2:O4").Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Cells(2, 16).Value = ws.Cells(maxpchangeindex, 9).Value  ' Ticker with max % increase
        ws.Cells(2, 17).Value = maxpchange
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = ws.Cells(minpchangeindex, 9).Value  ' Ticker with max % decrease
        ws.Cells(3, 17).Value = minpchange
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = ws.Cells(maxvolindex, 9).Value  ' Ticker with max volume
        ws.Cells(4, 17).Value = maxvol
        ws.Cells(4, 17).NumberFormat = "$#,##E+11"
        ' AutoFit columns for the current worksheet
        ws.Columns("I:Q").AutoFit
    Next ws
End Sub

