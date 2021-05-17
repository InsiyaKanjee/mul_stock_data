Sub uniqueticker()

 ' delcare worksheet loop variables
  Dim k, ws_count As Integer
  Dim ws As Worksheet
  ws_count = ActiveWorkbook.Worksheets.Count
  
  For k = 1 To ws_count
  
  Worksheets(k).Activate

      'declare loop variables for sorting through tickers
      Dim i, j As Integer

      ' Declare variable for holding the tickername
      Dim ticker As String
      
      ' Set an initial variable for holding the total volume
      Dim volume As Double
      volume = 0
      
      ' Keep track of the row location for each ticker
      Dim ticker_location As Long
      ticker_location = 2
      
      ' Count for how many iterations are run before non-unique value is hit
      Dim entrycount As Long
      entrycount = 0
      
      'keep track of opening price, closing price, yearly change in price & percentage change in price
      Dim openingprice, closingprice, pricechange, perchange As Double
      closingprice = 0
      openingprice = 0
      
      'identify last row of data
      Set sht = ActiveSheet
      Dim lastrow As Long
      
      'Refresh UsedRange
      sht.UsedRange
      lastrow = sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row
      
      'set headers
      Range("I1").Value = "Ticker"
      Range("I1").Columns.AutoFit
      Range("J1").Value = "Yearly Change"
      Range("J1").Columns.AutoFit
      Range("K1").Value = "Percentage Change"
      Range("K1").Columns.AutoFit
      Range("L1").Value = "Total Stock Volume"
      Range("L1").Columns.AutoFit
      
      
      ' Loop through all tickers
      For i = 2 To lastrow
        ' Check if we are still within the same ticker
            'not in same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
          ' Set & print ticker name
          ticker = Cells(i, 1).Value
          Range("I" & ticker_location).Value = ticker
          
          ' Set closing price
          closingprice = Cells(i, 6).Value
          
          ' calc & print total volume
          volume = volume + Cells(i, 7).Value
          Range("L" & ticker_location).Value = volume
          
          
          'calc & print price change
          pricechange = closingprice - openingprice
          Range("J" & ticker_location).Value = pricechange
          
          'conditional formating for price change
            If pricechange > 0 Then
                Range("J" & ticker_location).Interior.ColorIndex = 4
            Else
                Range("J" & ticker_location).Interior.ColorIndex = 3
            End If
          
          'calc & print percentage change
          If pricechange = 0 Then
            perchange = 0
          ElseIf openingprice = 0 Then
            Range("K" & ticker_location).Value = "exp inc in %"
          Else
            perchange = (pricechange) / openingprice * 100
          End If
          Range("K" & ticker_location).Value = perchange & "%"
                  
          ' increment ticker location so that next entry is added to the appropriate row
          ticker_location = ticker_location + 1
          
          ' Reset the volume & entry count
          volume = 0
          entrycount = 0
        
    
        Else
        ' If the cell immediately following a row is the same ticker
        
          ' Add to the volume & entry counter
          volume = volume + Cells(i, 7).Value
          entrycount = entrycount + 1
          
          'identify opening price
          openingprice = Cells(i - (entrycount - 1), 3).Value

        End If
    
      Next i
      
      'find max %inc, %dec & vol
      Dim maxdecvalue, maxincvalue, maxvolumevalue As Double
      Dim maxdecticker, maxincticker, maxvolticker As String
      maxdecvalue = 0
      maxincvalue = 0
      maxvolumevalue = 0
      
      For j = 2 To ticker_location - 1
      
        If Cells(j, 11).Value > maxincvalue Then
            maxincvalue = Cells(j, 11)
            maxincticker = Cells(j, 9)
        End If
     
        If Cells(j, 11) < maxdecvalue Then
            maxdecvalue = Cells(j, 11)
            maxdecticker = Cells(j, 9)
        End If
        
        If Cells(j, 12) > maxvolumevalue Then
            maxvolumevalue = Cells(j, 12)
            maxvolticker = Cells(j, 9)
        End If
        
      Next j
      
      ' print results
        Range("o2") = "Greatest % Increase"
        Range("o2").Columns.AutoFit
        Range("o3") = "Greatest % Decrease"
        Range("o3").Columns.AutoFit
        Range("o4") = "Greatest Total Volume"
        Range("o4").Columns.AutoFit
        
        Range("p1") = "Ticker"
        Range("p1").Columns.AutoFit
        Range("p2") = maxincticker
        Range("p3") = maxdecticker
        Range("p4") = maxvolticker
              
        
        Range("q1") = "Value"
        Range("q1").Columns.AutoFit
        Range("q2") = maxincvalue * 100 & "%"
        Range("q2").Columns.AutoFit
        Range("q3") = maxdecvalue * 100 & "%"
        Range("q3").Columns.AutoFit
        Range("q4") = maxvolumevalue
        Range("q4").Columns.AutoFit
        
    
    Next k
End Sub

