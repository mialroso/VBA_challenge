Sub multiYearStocks():

For Each ws In Worksheets
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   ws.Cells(1, 17).Value = "Value"
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   

Dim i As Long
Dim tickerName As String
Dim openPrice As Double
Dim totalVolume As Double
totalVolume = 0
Dim yearlyChange As Double
yearlyChange = 0
Dim percentChange As Double
Dim tickerRow As Long
tickerRow = 2
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
openPrice = ws.Cells(tickerRow, 3).Value

   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       tickerName = ws.Cells(i, 1).Value
       ws.Range("I" & tickerRow).Value = tickerName
   
       yearlyChange = yearlyChange + (ws.Cells(i, 6).Value - openPrice)
       ws.Range("J" & tickerRow).Value = yearlyChange
   
       percentChange = (yearlyChange / openPrice)
       ws.Range("K" & tickerRow).Value = percentChange
       ws.Range("K" & tickerRow).Style = "Percent"
       
       totalVolume = totalVolume + ws.Cells(i, 7).Value
       ws.Range("L" & tickerRow).Value = totalVolume
       
       
       tickerRow = tickerRow + 1
       yearlyChange = 0
       totalVolume = 0
       openPrice = ws.Cells(tickerRow, 3).Value
   Else
       totalVolume = totalVolume + ws.Cells(i, 7).Value
   End If
Next i


Dim yearLastRow As Long
yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To yearLastRow

   If ws.Cells(i, 10).Value >= 0 Then
       ws.Cells(i, 10).Interior.ColorIndex = 4
   Else
       ws.Cells(i, 10).Interior.ColorIndex = 3
   End If
Next i
   
Dim percentLastRow As Long
percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
Dim percent_max As Double
percent_max = 0
Dim percent_min As Double
percent_min = 0

For i = 2 To percentLastRow

   If percent_max < ws.Cells(i, 11).Value Then
       percent_max = ws.Cells(i, 11).Value
       ws.Cells(2, 17).Value = percent_max
       ws.Cells(2, 17).Style = "Percent"
       ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
   ElseIf percent_min > ws.Cells(i, 11).Value Then
       percent_min = ws.Cells(i, 11).Value
       ws.Cells(3, 17).Value = percent_min
       ws.Cells(3, 17).Style = "Percent"
       ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
   End If
Next i

Dim totalVolumeRow As Long
totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
Dim totalVolumeMax As Double
totalVolumeMax = 0

For i = 2 To totalVolumeRow

   If totalVolumeMax < ws.Cells(i, 12).Value Then
       totalVolumeMax = ws.Cells(i, 12).Value
       ws.Cells(4, 17).Value = totalVolumeMax
       ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
   End If
Next i
   
Next ws

End Sub




