Sub module2challenge():
  
'loop through worksheet
For Each ws In Worksheets
'Variables for Worksheet
Dim WorksheetName As String
Dim i As Long
Dim j As Long
'Variable for Ticker
Dim TickerCount As Long
'Variable for Last row of column A
Dim LastRowA As Long
'Variable for last row of column I
Dim LastRowI As Long
'Variable for percent change
Dim PercentChange As Double
'Variable for greatest increase
Dim GreatestIncrease As Double
'Variable for greatest decrease
Dim GreatestDecrease As Double
'Variable for greatest total volume
Dim GreatestVolume As Double
'Get the WorksheetName
WorksheetName = ws.Name
'Name headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
'Set Ticker Counter to first row
TickerCount = 2
'Set start row to 2
j = 2
'Find the last non-blank cell in column A
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox ("Last row in column A is " & LastRowA)
'Loop through all rows
For i = 2 To LastRowA
'Check if ticker name changed
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'Write ticker in column I
ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
'Calculate Yearly Change
ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
'If statement
If ws.Cells(TickerCount, 10).Value < 0 Then
'Set cell background color to red
ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
Else
'Set cell background color to green
ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
End If
'Calculate percent change
If ws.Cells(j, 3).Value <> 0 Then
PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
'Percent number formatting
ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
Else
ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
End If
'Calculate total volume
ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
'Add 1 to TickerCount
TickerCount = TickerCount + 1
'Set new start row of the ticker block
j = i + 1
End If
Next i
'Find last cell in column I
LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox ("Last row in column I is " & LastRowI)
GreatestVolume = ws.Cells(2, 12).Value
GreatestIncrease = ws.Cells(2, 11).Value
GreatestDecrease = ws.Cells(2, 11).Value
'Loop
For i = 2 To LastRowI
'Check if greatest volume is larger
If ws.Cells(i, 12).Value > GreatestVolume Then
GreatestVolume = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
Else
GreatestVolume = GreatestVolume
End If
'Check if greatest increase is larger
If ws.Cells(i, 11).Value > GreatestIncrease Then
GreatestIncrease = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
Else
GreatestIncrease = GreatestIncrease
End If
'Check if greatest decrease next value is smaller
If ws.Cells(i, 11).Value < GreatestDecrease Then
GreatestDecrease = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
Else
GreatestDecrease = GreatestDecrease
End If
'populate results in ws.Cells
ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
Next i
Worksheets(WorksheetName).Columns("A:Z").AutoFit
Next ws




End Sub
