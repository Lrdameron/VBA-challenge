Sub CalculateQuarterlyStockChangeWithSummaryPerSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim firstOpenPrice As Double
    Dim lastClosePrice As Double
    Dim totalVolume As Double ' Use Double to handle large numbers
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim i As Long
    Dim outputRow As Long
    Dim firstRow As Long ' To store the first row of each ticker for correct open price
    
    ' track greatest % increase, % decrease, and total volume for each sheet
    Dim greatestIncrease As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestVolume As Double
    Dim greatestVolumeTicker As String
    
    ' Loop through each sheet (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Sheets
        ' Initialize tracking variables for each sheet
        greatestIncrease = -99999
        greatestDecrease = 99999
        greatestVolume = 0
        
        ' Get the last row in the sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' columns I-L
        outputRow = 2 ' Starting row for output, assumes headers in row 1
        
        ' headers in columns I-L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' variables
        totalVolume = 0 ' Initialize total volume
        firstOpenPrice = ws.Cells(2, 3).Value ' Open price of the first entry (for first ticker)
        firstRow = 2 ' Initialize the first row for each ticker
        
        ' Loop through the rows of data
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            
            ' Accumulate volume for the current ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' If the next row is a new ticker or the last row, calculate the quarterly change
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                lastClosePrice = ws.Cells(i, 6).Value ' Close price of the last entry
                
                ' Calculate quarterly change
                quarterlyChange = lastClosePrice - firstOpenPrice
                If firstOpenPrice <> 0 Then
                    percentChange = quarterlyChange / firstOpenPrice ' No need to multiply by 100
                Else
                    percentChange = 0
                End If
                
                ' columns I-L for the current ticker
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = Round(percentChange, 4) ' Round to four decimal places for better precision
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' percent change in column K as a percentage
                ws.Cells(outputRow, 11).NumberFormat = "0.00%" ' Display as a percentage
                
                ' greatest % increase, % decrease, and total volume for each sheet
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                outputRow = outputRow + 1 ' Move to the next row for output
                
                ' Reset for the next ticker
                If i <> lastRow Then
                    firstOpenPrice = ws.Cells(i + 1, 3).Value ' Set open price for the next ticker
                    totalVolume = 0 ' Reset volume for the next ticker
                    firstRow = i + 1 ' Mark the first row of the new ticker
                End If
            End If
        Next i
        
        ' Output the greatest % increase, % decrease, and total volume in columns O-Q for the current sheet
        ' Write labels in P1 and Q1 in the current sheet
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Greatest % Increase
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = Round(greatestIncrease * 100, 2) & "%" ' Output percentage in Q
        
        ' Greatest % Decrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = Round(greatestDecrease * 100, 2) & "%" ' Output percentage in Q
        
        ' Greatest Total Volume
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = Format(greatestVolume, "0.00E+00") ' Output volume in Q
    Next ws
    
    MsgBox "Quarterly stock change calculations and summary are complete!"
End Sub

