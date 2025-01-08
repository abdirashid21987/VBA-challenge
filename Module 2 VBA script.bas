Attribute VB_Name = "Module1"
Sub StockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' Prepare the output headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 15).Value = "Metric"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ' Initialize variables
        outputRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        openingPrice = ws.Cells(2, 3).Value ' Initial opening price
        ' Loop through each row of data
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            ' Check for ticker change or last row
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                closingPrice = ws.Cells(i, 6).Value
                ' Calculate quarterly and percentage changes
                quarterlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentageChange = (quarterlyChange / openingPrice) * 100
                Else
                    percentageChange = 0
                End If
                ' Output results
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 10).NumberFormat = "0.00"
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalVolume
                ' Update greatest values
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    tickerGreatestIncrease = ticker
                End If
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    tickerGreatestDecrease = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                End If
                ' Reset for next ticker
                totalVolume = 0
                openingPrice = ws.Cells(i + 1, 3).Value
                outputRow = outputRow + 1
            End If
        Next i
        ' Output greatest values
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(2, 17).Value = tickerGreatestIncrease
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(3, 17).Value = tickerGreatestDecrease
        ws.Cells(4, 16).Value = greatestVolume
        ws.Cells(4, 17).Value = tickerGreatestVolume
        ' Conditional formatting for Quarterly Change
        With ws.Range("J2:J" & outputRow - 1)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"

.FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative
        End With
    Next ws
End Sub







