Sub stock_changes()

For Each ws In Worksheets

ws.Range("i1:l1").EntireColumn.Insert
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percentage change"
ws.Range("l1").Value = "Total stock volume"

Dim ticker_symbol As String

Dim yearly_change As Double
yearly_change = 0

Dim opening As Double
Dim closing As Double

Dim percentage As Double
percentage = 0
percentage = ws.Application.WorksheetFunction.RoundDown(percentage, 0)


Dim total_stock As Double
total_stock = 0

Dim summary_row As Integer
summary_row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

opening = ws.Cells(i, 3).Value
closing = ws.Cells(i, 6).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker_symbol = ws.Cells(i, 1).Value
    yearly_change = yearly_change + (closing - opening)
    percentage = yearly_change / opening
    total_stock = total_stock + (ws.Cells(i, 7).Value)

    ws.Range("i" & summary_row).Value = ticker_symbol
    ws.Range("j" & summary_row).Value = yearly_change
    ws.Range("k" & summary_row).Value = percentage
    ws.Range("l" & summary_row).Value = total_stock
    
    ws.Range("k" & summary_row).NumberFormat = "0.00%"

    summary_row = summary_row + 1
    yearly_change = 0
    percentage = 0
    total_stock = 0

Else
    total_stock = total_stock + (ws.Cells(i, 7).Value)
    yearly_change = yearly_change + (closing - opening)
    percentage = yearly_change / opening

End If

Next i

For i = 2 To lastrow

If ws.Cells(i, 10).Value > 0 Then

    ws.Cells(i, 10).Interior.Color = vbGreen

ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.Color = vbRed

Else
    ws.Cells(i, 10).Interior.ColorIndex = 0

End If

Next i



 Next ws
End Sub


