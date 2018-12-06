Attribute VB_Name = "Module1"
Sub test_stock_data()
'--------------------------------
  'Easy(summary ticker&total vol)
'--------------------------------
' LOOP THROUGH ALL SHEETS
 Dim ws As Worksheet
 For Each ws In Worksheets
     ws.Activate
  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String
  ' Set an initial variable for holding the total per ticker_vol
  Dim TickerVol_Total As Double
  TickerVol_Total = 0
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  ' Determine the Last Row
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through all rows
  For i = 2 To lastRow
  
  ' Check if we are still within the same ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the ticker name
      Ticker_Name = Cells(i, 1).Value
      ' Add to the TickerVol_Total
      TickerVol_Total = TickerVol_Total + Cells(i, 7).Value
      ' Print the ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name
      ' Print the ticker Vol Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = TickerVol_Total
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      ' Reset the Brand Total
      TickerVol_Total = 0
    ' If the cell immediately following a row is the same brand...
    Else
      ' Add to the TickerVol_Total
      TickerVol_Total = TickerVol_Total + Cells(i, 7).Value
    End If
  Next i
   'set column name for Summary Table
     Range("I1").Value = "Ticker"
     Range("J1").Value = "Total Stock Volume"
'--------------------------------
  'Moderate(yearly&percent change)
'--------------------------------

Dim openprice As Double
Dim closeprice As Double
Dim w, x, y, z As Long
Dim yearlychange, percentchange As Double

'Insert columns
Range("J1").EntireColumn.Insert
Range("J1").EntireColumn.Insert
'Add headers
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
'Determine lastrow of the dataset
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
'Set the first ticker's openprice
openprice = Range("C2").Value
' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row2 As Integer
Summary_Table_Row2 = 2
' Keep track of the location for the tickers that the openprice is equal to 0
y = 2

' Loop through all rows
For x = 2 To lastRow
    ' Searches for when the value of the next cell is different than that of the current cell
    If Range("A" & x + 1).Value <> Range("A" & x).Value Then
        closeprice = Range("F" & x).Value
        '--------------------------------------------------------------------------
         'by using the fllowing loop, the summary table can include all tickers
         'instead of exclude the tickers that never market opened.
        '--------------------------------------------------------------------------
        'Loop through the rows that the ticker's openprice for the firstday is 0
        If openprice = 0 Then
            'set the range of such tickers
            Set rng = Range(Range("C" & y), Range("C" & x))
            For z = y To x
            'Determine the row(j) that the openprice is not equal to 0 in rng
            If Range("C" & y) <> 0 Then
              j = Application.WorksheetFunction.Match(z, rng, 0)
              'Determine the row that theopen price is not equal to 0 in whole dataset
              openprice = Range("C" & y + j - 1).Value
              yearlychange = closeprice - openprice
              percentchange = yearlychange / openprice
              Range("J" & Summary_Table_Row2).Value = yearlychange
              Range("K" & Summary_Table_Row2).Value = Format(percentchange, "0.00%")
              
            Exit For
            Else
              yearlychange = 0
              percentchange = 0
              Range("J" & Summary_Table_Row2).Value = yearlychange
              Range("K" & Summary_Table_Row2).Value = percentchange
            End If
            Next z
            'look at the openprice in the next row
            openprice = Range("C" & x + 1).Value
            Summary_Table_Row2 = Summary_Table_Row2 + 1
            y = x + 1
        Else
            yearlychange = closeprice - openprice
            percentchange = yearlychange / openprice
            Range("J" & Summary_Table_Row2).Value = yearlychange
            Range("K" & Summary_Table_Row2).Value = Format(percentchange, "0.00%")
            'look at the openprice in the next row
            openprice = Range("C" & x + 1).Value
            Summary_Table_Row2 = Summary_Table_Row2 + 1
            y = x + 1
        End If
    End If
Next x

'Determine lastrow of the ticker in summary table
lastRow_ticker = Cells(Rows.Count, "I").End(xlUp).Row
'color the column of yearly change
For w = 2 To lastRow_ticker
   If Range("J" & w) > 0 Then
              Range("J" & w).Interior.ColorIndex = 4
              ElseIf Range("J" & w) < 0 Then
              Range("J" & w).Interior.ColorIndex = 3
              Else
              Range("J" & w).Interior.ColorIndex = 2
              End If
Next w

'--------------------------------
  'Hard
'--------------------------------
Dim rng1, rng2 As Range
Dim MAX_decrease, MAX_increase, MAX_volume As Double
Dim lastticker As Long
Dim a, b, c As Long

'set row&column name
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest Total Volumn"


'Set range from which to determine smallest value
lastticker = Cells(Rows.Count, "I").End(xlUp).Row
Set rng1 = Range("K1:K" & lastticker)
Set rng2 = Range("L1:L" & lastticker)

'Worksheet function MAX/MIN returns the value in a range
MAX_increase = Application.WorksheetFunction.Max(rng1)
MAX_decrease = Application.WorksheetFunction.Min(rng1)
MAX_volume = Application.WorksheetFunction.Max(rng2)

'Determine the rows that contain the above values
a = Application.WorksheetFunction.Match(MAX_increase, rng1, 0)
b = Application.WorksheetFunction.Match(MAX_decrease, rng1, 0)
c = Application.WorksheetFunction.Match(MAX_volume, rng2, 0)

'Displays value
Range("Q2").Value = Format(MAX_increase, "0.00%")
Range("Q3").Value = Format(MAX_decrease, "0.00%")
Range("Q4").Value = MAX_volume
Range("P2").Value = Range("I" & a)
Range("P3").Value = Range("I" & b)
Range("P4").Value = Range("I" & c)

Next ws

End Sub



