VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks():
'_________________________________________
'For Each Worksheet
'_________________________________________
For Each ws In Worksheets

'___________________________________
'Declare Variables
'___________________________________

'Ticker
Dim Ticker As String

'Variable to store Total Stock Amount
Dim StockTotal As Double
StockTotal = 0

'Build out Summary Table
Dim SumTableRow As Integer
SumTableRow = 2

'Variables to calculate Percent Change
Dim OpenAmount As Double
OpenAmount = 0
Dim CloseAmount As Double
CloseAmount = 0
Dim FirstStock As Double
FirstStock = 0
Dim YearlyChange As Double
YearlyChange = 0

'Variables to calculate greatest increase, decrease and Total Value
Dim Increase As Double
Increase = 0
Dim IncreaseRow As Double
IncreaseRow = 0
Dim Decrease As Double
Decrease = 0
Dim DecreaseRow As Double
DecreaseRow = 0
Dim TotalValue As Double
TotalValue = 0



'Handle unknown number of Rows
Dim LastRow As Long
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row

'Set Table Headers
ws.Range("K1").Value = "Ticker"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"
ws.Range("N1").Value = "Total Stock Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Value"

'Set Column Width for readability in columns K-R
ws.Columns("K:R").EntireColumn.AutoFit


'Set First Stock Value
OpenAmount = ws.Range("C2").Value

'Loop to Count Ticker Symbols
For i = 2 To LastRow

'___________________________________________
'Check if Ticker matches value above
'___________________________________________

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'Capture Ticker Name
Ticker = ws.Cells(i, 1).Value

'Calculate Stock Volume, and Close Price, Yearly Change
StockTotal = StockTotal + ws.Cells(i, 7).Value
CloseAmount = ws.Cells(i, 6).Value
YearlyChange = (CloseAmount - OpenAmount)

'Print Value to Summary Table
ws.Range("K" & SumTableRow).Value = Ticker
ws.Range("L" & SumTableRow).Value = YearlyChange
ws.Range("N" & SumTableRow).Value = StockTotal

If OpenAmount = 0 Then
ws.Range("M" & SumTableRow).NumberFormat = "0.00%"
ws.Range("M" & SumTableRow).Value = (CloseAmount - OpenAmount)

Else
ws.Range("M" & SumTableRow).NumberFormat = "0.00%"
ws.Range("M" & SumTableRow).Value = ((CloseAmount - OpenAmount) / OpenAmount)
End If

'__________________________________________________
'Format Cells by value - positive (Green)/ negative(red)
'_________________________________________________

If ws.Range("L" & SumTableRow).Value < 0 Then
ws.Range("L" & SumTableRow).Interior.ColorIndex = 3

Else
ws.Range("L" & SumTableRow).Interior.ColorIndex = 4
End If

'_________________________________________________
'Capture Greatest Total Value in summary table
'___________________________________________________
If StockTotal > ws.Range("R4").Value Then
ws.Range("R4").Value = StockTotal
ws.Range("Q4").Value = Ticker
End If
'________________________________________________

'Add a row to the summary table
SumTableRow = SumTableRow + 1

'Capture OpenAmount for next ticker
OpenAmount = ws.Cells(i + 1, 3).Value
'Reset Total Stock Volume
StockTotal = 0

'__________________________________________
'If Ticker Does match value below
'_________________________________________

Else
'Add a count to the numer if tickers
CountofTicker = (CountofTicker + 1)
'Sum Stock Volume While Ticker matches
StockTotal = StockTotal + ws.Cells(i, 7).Value
End If

Next i
'_______________________________________________________
'Find Greatest % Increase and Decrease, print values
'_______________________________________________________
Increase = Application.WorksheetFunction.Max(ws.Columns("M").Value)
ws.Range("R2").Value = Increase
ws.Range("R2").NumberFormat = "0.00%"
IncreaseRow = Application.WorksheetFunction.Match(Increase, ws.Range("M:M"), 0)
ws.Range("Q2").Value = ws.Range("K" & IncreaseRow).Value

Decrease = Application.WorksheetFunction.Min(ws.Columns("M").Value)
ws.Range("R3").Value = Decrease
ws.Range("R3").NumberFormat = "0.00%"
DecreaseRow = Application.WorksheetFunction.Match(Decrease, ws.Range("M:M"), 0)
ws.Range("Q3").Value = ws.Range("K" & DecreaseRow).Value

Next ws

End Sub

