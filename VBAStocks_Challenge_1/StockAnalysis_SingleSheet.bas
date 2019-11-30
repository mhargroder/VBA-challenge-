Attribute VB_Name = "StockAnalysis_SingleSheet"

Sub StockAnalysis_SingleSheet()
'm.hargroder@outlook.com
'This module includes the functionality for easy and moderate requirements of the homework challenge
'A second module is also provided in VBAChallenge2 dir that provides wrapper loop to do for each worksheet in a workbook

Dim firstOpen As Double, lastClose As Double, sumVol As Double
Dim sumRow As Integer
     
 sumVol = 0
 sumRow = 1
 dataRows = Range("A" & Rows.Count).End(xlUp).Row

      For j = 2 To dataRows 'build the summary range by evaluating each days trading activity for a ticker and record results
          
          'Since most common check if j row is not a first or last for a ticker
           If Cells(j, 1).Value = Cells(j - 1, 1).Value And Cells(j, 1).Value = Cells(j + 1, 1).Value Then
           sumVol = sumVol + Cells(j, 7).Value
                  
            ' Check if j row is first row of a ticker
            ElseIf Cells(j, 1).Value <> Cells(j - 1, 1).Value And Cells(j, 1).Value = Cells(j + 1, 1).Value Then
             firstOpen = Cells(j, 3).Value
             sumVol = Cells(j, 7).Value
                       
             'Check if j row is last row of a ticker
             ElseIf Cells(j, 1).Value = Cells(j - 1, 1).Value And Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
              sumRow = sumRow + 1
              lastClose = Cells(j, 6) 'doing this for readability..could have just used cell ref in this block and no var necessary
              Cells(sumRow, 9).Value = Cells(j, 1).Value 'record the ticker
              Cells(sumRow, 10).Value = lastClose - firstOpen 'record the delta
               If lastClose > 0 Then
                Cells(sumRow, 11).Value = (lastClose - firstOpen) / lastClose 'record % change
               Else
                Cells(sumRow, 11).Value = 0
               End If
              Cells(sumRow, 12).Value = sumVol + Cells(j, 7).Value 'record the sum volume

           End If
     
         Next j

         'Format the summary range
         
          ' Set summary headers
          Range("I1").Value = "Ticker"
          Range("J1").Value = "YearlyChange"
          Range("K1").Value = "PercentChange"
          Range("L1").Value = "TotalStockVolume"
 
         
          'Conditionally format the  yearly change cells
          For k = 2 To sumRow
           If Cells(k, 10) < 0 Then
            Cells(k, 10).Interior.Color = RGB(255, 0, 0)
           Else
            Cells(k, 10).Interior.Color = RGB(0, 255, 0)
           End If
          Next k
          
          'Format the percent change cells
          Range("K2", "K" & sumRow).Style = "Percent"
          
          'Set up the Greatest Increase and Decrease Section
           'Set headers
          Range("P2").Value = "Ticker"
          Range("Q2").Value = "Value"
          Range("O3").Value = "Greatest%Increase"
          Range("O4").Value = "Greatest%Decrease"
          Range("O5").Value = "GreatestTotalVolume"
          
           'Format percent cells
          Range("Q3:Q4").Style = "Percent"
           'Resize the cell widths
           Columns("M:N").ColumnWidth = 6#
           Columns("Q:Q").ColumnWidth = 11.2
           Columns("O:P").AutoFit
                     
        'search the range once with function and set to var
    
         Dim mxInc As Double, mxDcr As Double, mxVol As Double
         mxInc = WorksheetFunction.Max(Range("K3" & ":K" & sumRow))
         mxDcr = WorksheetFunction.Min(Range("K3" & ":K" & sumRow))
         mxVol = WorksheetFunction.Max(Range("L3" & ":L" & sumRow))
         

          'Loop through to find the tickers for the max inc,dec, vol
           For p = 2 To sumRow
            If Cells(p, 11) = mxInc Then
             Cells(3, 16).Value = Cells(p, 9).Value
             Cells(3, 17).Value = Cells(p, 11).Value
             
              ElseIf Cells(p, 11) = mxDcr Then
               Cells(4, 16).Value = Cells(p, 9).Value
                Cells(4, 17).Value = Cells(p, 11).Value
                
                 ElseIf Cells(p, 12) = mxVol Then
                  Cells(5, 16).Value = Cells(p, 9).Value
                  Cells(5, 17).Value = Cells(p, 12).Value
                
             End If
            Next p

End Sub
