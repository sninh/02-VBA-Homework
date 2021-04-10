Sub stocks()

Dim ws As Worksheet

For Each ws In Worksheets

'add headers for the new columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "% Change"
ws.Range("L1").Value = "Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
Dim i As Long
Dim yearly As Double
Dim percent As Double
Dim first As Double
Dim last As Double
Dim a As Long
Dim increase As Double
Dim decrease As Double
Dim bigvolume As Double
Dim increaseticker As String
Dim decreaseticker As String
Dim bigvolumeticker As String



  ' Set an initial variable for holding the ticker
  Dim ticker_Name As String
  ' Set an initial variable for holding the total volume
  Dim volume_Total As Double
  volume_Total = 0
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Long
  'summary table starts on row 2
  Summary_Table_Row = 2
  'first price of ticker is c2
first = ws.Range("C2")
  'used to put values in summary table
  a = 2
  ' Loop through all tickers
  For i = 2 To ws.Range("A1").End(xlDown).Row
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ' Set the Brand name
      ticker_Name = ws.Cells(i, 1).Value
      ' Add to the Brand Total
      volume_Total = volume_Total + ws.Range("G" & i).Value
      ' Print the ticker name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker_Name
      ' Print the volume amount the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = volume_Total
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      ' Reset the Brand Total
      volume_Total = 0
      'last value for the ticker for the year
      last = ws.Cells(i, 6).Value
      
      'calculate the yearly amount divided by the first price if 0
      If first = 0 Then
      yearly = 0
      Else: yearly = last - first
      End If:
      
      
      'put yearly value in column J next to ticker
      ws.Cells(a, 10).Value = yearly
      'calculate the percent amount divided by the first price if 0
      If first = 0 Then
      percent = 0
      Else: percent = yearly / first
      End If
    
      'put percent in column k next to ticker
      ws.Cells(a, 11).Value = percent
      
      If ws.Cells(Summary_Table_Row - 1, 10).Value > 0 Then
            'if cell is more than 0 make it green
            ws.Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(Summary_Table_Row - 1, 10).Value < 0 Then
            'if cell is less than 0 make it red
            ws.Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 3
            Else
                'if there is nothing in the cell, don't put a color
            ws.Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 0
            End If
            'make percent change a %
            ws.Cells(Summary_Table_Row - 1, 11).NumberFormat = "0.00%"
            
      'next row for tickers
      a = a + 1
      'have the first price be the last row for the previous ticker + 1
      first = ws.Cells(i + 1, 3).Value
                
      'Conditional formatting: positive in green & negative in red
            

    Else
      ' Add to the volume Total
      volume_Total = volume_Total + ws.Range("G" & i).Value
    End If
    
  Next i
  
Dim biggest As Double
Dim maxticker As String
biggest = Application.WorksheetFunction.max(ws.Range("K:K"))
ws.Range("P2").Value = biggest
maxticker = (WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("P2").Value, ws.Range("K:K"), 0)))
ws.Range("O2").Value = maxticker
ws.Range("P2").NumberFormat = "0.00%"

Dim smallest As Double
Dim minticker As String
smallest = Application.WorksheetFunction.min(ws.Range("K:K"))
ws.Range("P3").Value = smallest
minticker = (WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("P3").Value, ws.Range("K:K"), 0)))
ws.Range("O3").Value = minticker
ws.Range("P3").NumberFormat = "0.00%"

Dim maxvolumeticker As String
Dim maxvolume As Double
maxvolume = Application.WorksheetFunction.max(ws.Range("L:L"))
ws.Range("P4").Value = maxvolume
maxvolumeticker = (WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("P4").Value, ws.Range("L:L"), 0)))
ws.Range("O4").Value = maxvolumeticker

Next ws

End Sub

