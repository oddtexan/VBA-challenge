Sub stocksummary()

  Dim Current As Worksheet
  
  For Each Current In ThisWorkbook.Worksheets
  Current.Activate

    ' Set an initial variable for holding the TICKER name
    Dim Ticker As String

    ' Set an initial variable for holding the TOTAL PER TICKER
    Dim total_volume As Double
    total_volume = 0

    ' Keep track of the location for each ticker in the SUMMARY TABLE
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    ' Set the header of SUMMARY TABLE
    Range("I1") = "Ticker"
    Range("J1") = "Quarterly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % Increase Value"
    Range("O3") = "Greatest % Decrease Value"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
  
    'Set the LAST ROW (from lecture #2 credit card example)
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Loop through all daily records
       For i = 2 To lastrow
        'If previous TICKER and current ticker are not the same, then...
        If Cells(i - 1, 1) <> Cells(i, 1) Then
        opening_price = Cells(i, 3)
        
        'If next TICKER and CURRENT TICKER are not the same, then...
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the TICKER name
        Ticker = Cells(i, 1).Value

        ' Add to the TOTAL VOLUME
        total_volume = total_volume + Cells(i, 7).Value
      
        'Set the CLOSING PRICE
        closing_price = Cells(i, 6).Value
      
        'Calculate the change between OPENING PRICE and CLOSING PRICE
        quarterly_change = closing_price - opening_price
      
        'Calculate the PERCENT CHANGE between OPENING PRICE and CLOSING PRICE
        'If opening_price = 0 Then
        'percentage_change = Null
        percent_change = ((closing_price - opening_price) / opening_price)
        On Error Resume Next

        ' Print the TICKER in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
      
        ' Print the Quarterly CHANGE to the Summary Table
        Range("J" & Summary_Table_Row).Value = quarterly_change
      
        ' Print the QUARTERLY CHANGE to the Summary Table
        Range("K" & Summary_Table_Row).Value = percent_change
        Columns("K:K").NumberFormat = "0.00%"

        ' Print the TICKER AMOUNT to the Summary Table
        Range("L" & Summary_Table_Row).Value = total_volume
      
        ' Add one to the SUMMARY TABLE row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the TOTAL VOLUME
        total_volume = 0

        ' If TICKER is the same
        Else
        'Add to the TICKER TOTAL
        total_volume = total_volume + Cells(i, 7).Value
        End If
             
       Next i
        
        'After the 1st loop is done, set the NEXT loop
       Dim greatest_increase, greatest_decrease As Double
       greatest_increase = Cells(2, 11)
       greatest_decrease = Cells(2, 11)
       greatest_volume = Cells(2, 12)
       lastrow_summary = Cells(Rows.Count, 10).End(xlUp).Row
  
       For j = 2 To lastrow_summary
        'Change the FORMAT depending on the value
        If Cells(j, 10) >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
  
        ElseIf Cells(j, 10) < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        'Loop through each row and replace the GREATEST INCREASE VALUE
        If Cells(j, 11) > greatest_increase Then
        greatest_increase = Cells(j, 11)
        Cells(2, 17) = greatest_increase
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16) = Cells(j, 9)
        End If
   
        'Loop through each row and replace the GREATEST DECREASE VALUE
        If Cells(j, 11) < greatest_decrease Then
        greatest_decrease = Cells(j, 11)
        Cells(3, 17) = greatest_decrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16) = Cells(j, 9)
        End If
        
        'Loop through each row and replace the GREATEST VOLUME TOTAL
        If Cells(j, 12) > greatest_volume Then
        greatest_volume = Cells(j, 12)
        Cells(4, 17) = greatest_volume
        Cells(4, 16) = Cells(j, 9)
        End If
   
       Next j
 
        Columns("I:Q").AutoFit
        
 Next

End Sub


