VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MULTIPLE_YEAR_STOCK_DATA()
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
  ws.Activate
  Debug.Print ws.Name
  

    ' Set an initial variable for holding the ticker name
    Dim ticker_name As String

    ' Set an initial variable for holding the total per ticker
    Dim total_volume As Double
    total_volume = 0

    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
 
    ' Set the header of summary table
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Quarterly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
 
    'Set the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    ' Loop through all daily records
       For i = 2 To lastrow
        'If previous ticker and current ticker are not the same, then...
        If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
        opening_price = ws.Cells(i, 3)
       
        'If next ticker and current ticker are not the same, then...
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Set the ticker name
        ticker_name = ws.Cells(i, 1).Value

        ' Add to the total volume
        total_volume = total_volume + ws.Cells(i, 7).Value
     
        'Set the closing price
        closing_price = ws.Cells(i, 6).Value
     
        'Calculate the change between opening price and closing price
        quarterly_change = closing_price - opening_price
     
        'Calculate the percent change between opening price and closing price
        'If opening_price = 0 Then
        'percentage_change = Null
        percent_change = ((closing_price - opening_price) / opening_price)
        On Error Resume Next

        ' Print the ticker in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker_name
     
        ' Print the yearly change to the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = quarterly_change
     
        ' Print the yearly change to the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        ws.Columns("K:K").NumberFormat = "0.00%"

        ' Print the ticker Amount to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = total_volume
     
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
     
        ' Reset the total volume
        total_volume = 0

        ' If ticker is the same
        Else
        'Add to the ticker total
        total_volume = total_volume + ws.Cells(i, 7).Value

        End If
             
       Next i
       
        'After the 1st loop is done, set the next loop
       Dim greatest_increase, greatest_decrease As Double
       greatest_increase = ws.Cells(2, 11)
       greatest_decrease = ws.Cells(2, 11)
       greatest_volume = ws.Cells(2, 12)
       lastrow_summary = ws.Cells(Rows.Count, 10).End(xlUp).Row
 
       For j = 2 To lastrow_summary
        'Change the format depending on the value
        If ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
 
        ElseIf ws.Cells(j, 10) < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
   
        End If
       
       If ws.Cells(j, 11) > 0 Then
       ws.Cells(j, 11).Interior.ColorIndex = 4
       
       ElseIf ws.Cells(j, 11) < 0 Then
       ws.Cells(j, 11).Interior.ColorIndex = 3
       
       End If
       
        'Loop through each row and replace the greatest increase value
        If ws.Cells(j, 11) > greatest_increase Then
        greatest_increase = ws.Cells(j, 11)
        ws.Cells(2, 17) = greatest_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16) = ws.Cells(j, 9)
   
        End If
   
        'Loop through each row and replace the greatest decrease value
        If ws.Cells(j, 11) < greatest_decrease Then
        greatest_decrease = ws.Cells(j, 11)
        ws.Cells(3, 17) = greatest_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16) = ws.Cells(j, 9)
 
        End If
       
        'Loop through each row and replace the greatest total volume
        If ws.Cells(j, 12) > greatest_volume Then
        greatest_volume = ws.Cells(j, 12)
        ws.Cells(4, 17) = greatest_volume
        ws.Cells(4, 16) = ws.Cells(j, 9)
   
        End If
   
       Next j
    Next
        ws.Columns("I:Q").AutoFit
 
End Sub
