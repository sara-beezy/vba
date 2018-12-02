Sub multi_stock_2014()

  ' Set an initial variable for holding the brand name
  Dim stockName As String
  Dim c As Long
  Dim s1, s2, s3 As Worksheet
  
  Set s1 = Sheets("2014")

  
  ' Set an initial variable for holding the total
  Dim stockTotal As Double
  stockTotal = 0

  ' Keep track of the location for each stock
  Dim StockTableRow As Integer
StockTableRow = 2

s1.Activate
Range("J1").Value = "Ticker"
Range("K1").Value = "Total Stock Volume"

easy = WorksheetFunction.CountA(Range("A:A"))

  ' Loop through all stocks

  For c = 2 To easy

        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then

      ' Set the Brand name
      stockName = Cells(c, 1).Value

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(c, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("J" & StockTableRow).Value = stockName

      ' Print the Brand Amount to the Summary Table
      Range("K" & StockTableRow).Value = stockTotal

      ' Add one to the summary table row
     StockTableRow = StockTableRow + 1
      
      ' Reset the Brand Total
      stockTotal = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(c, 7).Value

    End If

  Next c

End Sub

Sub multi_stock_2015()

  ' Set an initial variable for holding the brand name
  Dim stockName As String
  Dim c As Long
  Dim s1, s2, s3 As Worksheet
  
  Set s2 = Sheets("2015")

  
  ' Set an initial variable for holding the total
  Dim stockTotal As Double
  stockTotal = 0

  ' Keep track of the location for each stock
  Dim StockTableRow As Integer
StockTableRow = 2

s2.Activate
Range("J1").Value = "Ticker"
Range("K1").Value = "Total Stock Volume"

easy = WorksheetFunction.CountA(Range("A:A"))

  ' Loop through all stocks

  For c = 2 To easy

        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then

      ' Set the Brand name
      stockName = Cells(c, 1).Value

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(c, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("J" & StockTableRow).Value = stockName

      ' Print the Brand Amount to the Summary Table
      Range("K" & StockTableRow).Value = stockTotal

      ' Add one to the summary table row
     StockTableRow = StockTableRow + 1
      
      ' Reset the Brand Total
      stockTotal = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(c, 7).Value

    End If

  Next c

End Sub
Sub multi_stock_2016()

  ' Set an initial variable for holding the brand name
  Dim stockName As String
  Dim c As Long
  Dim s1, s2, s3 As Worksheet
  
  Set s3 = Sheets("2016")

  
  ' Set an initial variable for holding the total
  Dim stockTotal As Double
  stockTotal = 0

  ' Keep track of the location for each stock
  Dim StockTableRow As Integer
StockTableRow = 2

s3.Activate
Range("J1").Value = "Ticker"
Range("K1").Value = "Total Stock Volume"


easy = WorksheetFunction.CountA(Range("A:A"))

  ' Loop through all stocks

  For c = 2 To easy

        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then

      ' Set the Brand name
      stockName = Cells(c, 1).Value

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(c, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("J" & StockTableRow).Value = stockName

      ' Print the Brand Amount to the Summary Table
      Range("K" & StockTableRow).Value = stockTotal

      ' Add one to the summary table row
     StockTableRow = StockTableRow + 1
      
      ' Reset the Brand Total
      stockTotal = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      stockTotal = stockTotal + Cells(c, 7).Value

    End If

  Next c

End Sub