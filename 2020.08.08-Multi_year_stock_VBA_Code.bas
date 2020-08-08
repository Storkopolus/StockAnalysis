Attribute VB_Name = "Module1"
Sub StockSummary()

' Set a variable for specifying the column of interest
  Dim ws As Worksheet
   
For Each ws In Worksheets
  Dim column As Double
  column = 1
  Dim InitialStockCost0 As Double
  Dim FinalStockCost As Double
'Stock Volume is the total number of trades
  Dim StockVolume As Double
'Total Stocks is the number of stocks with different ticker symbols that have been cycled through
  Dim TotalStocks As Double
  TotalStocks = 1
  Dim LastRow As Double
  Dim CurrentStock As String
  Dim YearlyChange As Double
  Dim PercentChange As Double

' Add required Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

  
' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Sets initial Stock Cost for First Ticker Symbol
    InitialStockCost0 = ws.Cells(2, 3).Value

  
  
' Loop through rows in the column
    For i = 2 To LastRow
          
' Searches for when the value of the next cell is equal to the current cell
    If ws.Cells(i + 1, column).Value = ws.Cells(i, column).Value Then
    StockVolume = StockVolume + ws.Cells(i, 7).Value
    CurrentStock = ws.Cells(i, 1).Value
        
' Searches for when the value of the next cell is different than that of the current cell
    ElseIf ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then

' Setting variables based on a Stock Ticker Company Change - 1
    FinalStockCost = ws.Cells(i, 6).Value
    ws.Cells(TotalStocks + 1, 9).Value = CurrentStock
    YearlyChange = FinalStockCost - InitialStockCost0
    ws.Cells(TotalStocks + 1, 10).Value = YearlyChange
    
'division by 0 workaround
    If (InitialStockCost0 = 0 And FinalStockCost = 0) Then
    PercentChange = 0
    ElseIf (InitialStockCost0 = 0 And FinalStockCost <> 0) Then
    PercentChange = 1
    Else
'divide closing price by open price
    PercentChange = YearlyChange / InitialStockCost0
    End If
    ws.Cells(TotalStocks + 1, 11).Value = FormatPercent(PercentChange)
    ws.Cells(TotalStocks + 1, 12).Value = StockVolume + ws.Cells(i, 7)
        
'If Statement to color the YearlyChange cell
    If YearlyChange < 0 Then
    ws.Cells(TotalStocks + 1, 10).Interior.ColorIndex = 3
    ElseIf YearlyChange > 0 Then
    ws.Cells(TotalStocks + 1, 10).Interior.ColorIndex = 4
    End If
    
' Adding one to the TotalStocks count
    TotalStocks = TotalStocks + 1
    
'Finds Inital Cost of the Upcoming Stock
    InitialStockCost0 = ws.Cells((i + 1), 3).Value
    
'Reset StockVolume to 0
    StockVolume = 0
    
    End If

  Next i
  
Next ws

End Sub
