VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AnalyzeStocks()
    ' Declare variables
    Dim i As Long ' Row number
    Dim cell_vol As LongPtr ' Volume of stock traded in a cell
    Dim vol_total As LongPtr ' Total volume of the stock
    Dim ticker As String ' Stock symbol
    Dim k As Long ' Leaderboard row
    Dim ticker_close As Double ' Closing price of a stock
    Dim ticker_open As Double ' Opening price of a stock
    Dim price_change As Double ' Change in stock price
    Dim percent_change As Double ' Percentage change in stock price
    Dim greatest_increase As Double ' Greatest percentage increase
    Dim greatest_decrease As Double ' Greatest percentage decrease
    Dim greatest_volume As LongPtr ' Greatest total volume
    Dim greatest_increase_ticker As String ' Ticker symbol for greatest percentage increase
    Dim greatest_decrease_ticker As String ' Ticker symbol for greatest percentage decrease
    Dim greatest_volume_ticker As String ' Ticker symbol for greatest total volume
    Dim lastRow As Long ' Last row of data
    Dim ws As Worksheet ' Worksheet iterator

    ' Initialize variables
    greatest_increase = -999999999 ' Set initial values to very low/high numbers
    greatest_decrease = 999999999
    greatest_volume = 0

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Initialize variables
        vol_total = 0 ' Starting with zero total volume
        k = 2 ' Starting the leaderboard from row 2

        ' Set up the leaderboard headers
        ws.Range("I1:L1").Value = Array("Ticker", "Quarterly Change", "Percent Change", "Volume")
        ws.Range("I1:L1").Font.Bold = True

        ' Assign the opening price for the first stock
        ticker_open = ws.Cells(2, 3).Value ' Assuming the opening price is in column C

        ' Loop through each row of data
        For i = 2 To lastRow
            cell_vol = ws.Cells(i, 7).Value ' Assuming the volume is in column G
            ticker = ws.Cells(i, 1).Value ' Assuming the ticker symbol is in column A
            
            ' Check if the next row's ticker is different
            If ws.Cells(i + 1, 1).Value <> ticker Then
                ' Different ticker
                vol_total = vol_total + cell_vol
                ' Get the closing price of the stock
                ticker_close = ws.Cells(i, 6).Value ' Assuming the closing price is in column F
                price_change = ticker_close - ticker_open
                
                ' Calculate percentage change, protecting against division by zero
                If ticker_open <> 0 Then
                    percent_change = price_change / ticker_open
                Else
                    percent_change = 0
                End If
                
                ' Output data to the leaderboard
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = price_change
                ws.Cells(k, 11).Value = percent_change
                ws.Cells(k, 12).Value = vol_total
                
                ' Apply conditional formatting for percent change
                If percent_change > 0 Then
                    ws.Cells(k, 11).Interior.ColorIndex = 4 ' Green
                ElseIf percent_change < 0 Then
                    ws.Cells(k, 11).Interior.ColorIndex = 3 ' Red
                End If
                
                ' Update extreme values
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
                End If
                
                If vol_total > greatest_volume Then
                    greatest_volume = vol_total
                    greatest_volume_ticker = ticker
                End If
                
                ' Reset variables for the next stock
                vol_total = 0
                k = k + 1
                ticker_open = ws.Cells(i + 1, 3).Value ' Move to the next row for the opening price
            Else
                ' Add to total volume
                vol_total = vol_total + cell_vol
            End If
        Next i
        
        ' Apply number format to the percent change column
        ws.Columns("K:K").NumberFormat = "0.00%"
        
        ' Autofit columns for a better look
        ws.Columns("I:L").AutoFit
        
        ' Output the results for greatest percentage increase, decrease, and total volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(2, 15).Value = greatest_increase_ticker
        ws.Cells(3, 15).Value = greatest_decrease_ticker
        ws.Cells(4, 15).Value = greatest_volume_ticker
        
        ws.Cells(2, 16).Value = Format(greatest_increase, "0.00%")
        ws.Cells(3, 16).Value = Format(greatest_decrease, "0.00%")
        ws.Cells(4, 16).Value = Format(greatest_volume, "#,##0")
    Next ws
End Sub


