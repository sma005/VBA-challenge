Attribute VB_Name = "Module1"
Sub AlphabeticStockQuery()

'This code takes an Excel file full of stocks, and lists their yearly change, their percent change, and their total volume for the year.
'Each sheet has to have stocks with the same name grouped together
'Not sure why I can't assign Array2 to Array1 example: currentStock = previousStock
'In my opinion, this should be an Array, but Arrays with a Variant datatype seem to not be able to be assigned

'Dim numOfSheets As Integer
'numOfSheets = ThisWorkbook.Worksheets.Count

Dim numOfRows As Long

'aggregate declarations
Dim i As Long

'ticker only needs to hold Name,open,close,volume,totalVolume
'Different datatypes than others learned from class are found here...
'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

Dim currentStockName As String
Dim currentStockOpen As Single
Dim currentStockClose As Single
Dim currentStockVolume As Double
Dim currentStockTotalVolume As Double

Dim firstDayStockName As String
Dim firstDayStockOpen As Single
Dim firstDayStockVolume As Double
Dim firstDayStockTotalVolume As Double 'might be able to take this out

Dim previousStockName As String
Dim previousStockOpen As Single
Dim previousStockClose As Single
Dim previousStockVolume As Double
Dim previousStockTotalVolume As Double

Dim lastDayStockName As String
Dim lastDayStockClose As Single
Dim lastDayStockVolume As Double
Dim lastDayStockTotalVolume As Double 'might be able to take this out

Dim outputRow As Integer

'Start
For Each ws In Worksheets

ws.Cells(1, 9) = "Stock"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

outputRow = 2 'First line to output data

numOfRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'reset lastDayStock on new sheet
lastDayStockName = "blank_value"
lastDayStockClose = 0
lastDayStockVolume = 0
lastDayStockTotalVolume = 0

'Headers for Aggregates
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
 
   For i = 2 To numOfRows 'first row is header
    'grab row
    currentStockName = ws.Cells(i, 1) 'name
    currentStockOpen = ws.Cells(i, 3) 'open
    currentStockClose = ws.Cells(i, 6) 'close
    currentStockVolume = ws.Cells(i, 7) 'volume
    
    If (lastDayStockName <> previousStockName) Then
        lastDayStockTotalVolume = 0
    End If

    'generate total column and make sure they are the same ticker
    If (currentStockName = previousStockName) Then
       currentStockTotalVolume = previousStockTotalVolume + currentStockTotalVolume 'running total

      'make current day last day
      lastDayStockName = currentStockName
      lastDayStockClose = currentStockClose
      lastDayStockVolume = currentStockVolume
      lastDayStockTotalVolume = currentStockVolume + lastDayStockTotalVolume
      
    'different ticker - reset first day ticker
    Else
            If (lastDayStockName <> "blank_value") Then
            
                  'print previous stock results
                   ws.Cells(outputRow, 9) = lastDayStockName 'previousStockName
            
                   'Need to find Range from cell number to apply formatting
                   'https://stackoverflow.com/questions/31986047/how-to-use-a-cell-value-as-part-of-a-range-using-vba
                    ws.Cells(outputRow, 10) = lastDayStockClose - firstDayStockOpen
                       If (ws.Cells(outputRow, 10) < 0) Then
                          ws.Range("J" & outputRow).Interior.ColorIndex = 3
                       Else
                          ws.Range("J" & outputRow).Interior.ColorIndex = 4
                          
                       End If
                  
                  ws.Range("J" & outputRow).NumberFormat = "0.00"
                  ws.Cells(outputRow, 11) = (lastDayStockClose - firstDayStockOpen) / firstDayStockOpen
                  ws.Range("K" & outputRow).NumberFormat = "0.00%"
                  
                  ws.Cells(outputRow, 12) = lastDayStockTotalVolume 'test
                  
           outputRow = outputRow + 1
           
           End If
            
            'new stock - reset aggregate
            firstDayStockName = currentStockName
            firstDayStockOpen = currentStockOpen
            firstDayStockVolume = currentStockVolume
            firstDayStockTotalVolume = 0

    End If
    
        previousStockName = currentStockName
        previousStockOpen = currentStockOpen
        previousStockClose = currentStockClose
        previousStockVolume = currentStockVolume
        previousStockTotalVolume = currentStockTotalVolume
    
    Next i
    
  'https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba
  'https://www.statology.org/vba-xlookup/
  'Max Increase
  ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
  ws.Range("P2").Value = WorksheetFunction.XLookup(ws.Range("Q2"), ws.Range("K:K"), ws.Range("I:I"))
  ws.Range("Q2").NumberFormat = "0.00%"
  
  'Min Increase
  ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
  ws.Range("P3").Value = WorksheetFunction.XLookup(ws.Range("Q3"), ws.Range("K:K"), ws.Range("I:I"))
  ws.Range("Q3").NumberFormat = "0.00%"
  
  'Greatest Total Volume
  ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
  ws.Range("P4").Value = WorksheetFunction.XLookup(ws.Range("Q4"), ws.Range("L:L"), ws.Range("I:I"))
  
Next ws

End Sub

