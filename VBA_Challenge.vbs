Sub VBA_Challenge()


'Loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets


'Add column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


'Determine last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Declare variables
Dim Ticker_Name As String
Dim Total_Volume, StartPrice, EndPrice, YearlyChange, PercentChange As Double
  Total_Volume = 0
Dim WorksheetName As String
  WorksheetName = ws.Name
'Location for information in the summary table
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2


'Loop through all tickers
For i = 2 To LastRow

  
  If ws.Cells(i, 2).Value = ws.Name + "0102" Then
    StartPrice = ws.Cells(i, 3).Value
  End If
    
  
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker_Name = ws.Cells(i, 1).Value
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    EndPrice = ws.Cells(i, 6).Value
    
    
    'Calculate the Yearly Change and Percent Change
    YearlyChange = EndPrice - StartPrice
    If StartPrice <> 0 Then
      PercentChange = (YearlyChange / StartPrice)
    Else
      PercentChange = 0
    End If
    
    
    'Print the information in the Summary Table
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    ws.Range("L" & Summary_Table_Row).Value = Total_Volume
    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
    ws.Range("K" & Summary_Table_Row).Value = PercentChange
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     
     
    'Apply conditional formatting
    If YearlyChange > 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
    If PercentChange > 0 Then
      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    End If


    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
      
      
    'Reset the Total Volume
    Total_Volume = 0


  'If the cell in the row below is the same ticker:
  Else
    'Add to the Total Volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
  End If

 Next i
 
 
'Determine last row of summary table
 LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row


'Declare variables
 Dim GPI_Ticker, GPD_Ticker, GTV_Ticker As String
 Dim GPI_Value, GPD_Value, GTV_Value As Double
  GPI_Ticker = ws.Cells(2, 9).Value
  GPD_Ticker = ws.Cells(2, 9).Value
  GTV_Ticker = ws.Cells(2, 9).Value
  GPI_Value = ws.Cells(2, 11).Value
  GPD_Value = ws.Cells(2, 11).Value
  GTV_Value = ws.Cells(2, 12).Value
 
 
'Loop through summary table
 For i = 3 To LastRow
 
 
  'Greatest Percent Increase
  If ws.Cells(i, 11) > GPI_Value Then
    GPI_Value = ws.Cells(i, 11).Value
    GPI_Ticker = ws.Cells(i, 9).Value
  End If
  
  
  'Greatest Percent Decrease
  If ws.Cells(i, 11) < GPD_Value Then
    GPD_Value = ws.Cells(i, 11).Value
    GPD_Ticker = ws.Cells(i, 9).Value
  End If
  
  
  'Greatest Total Volume
  If ws.Cells(i, 12) > GTV_Value Then
    GTV_Value = ws.Cells(i, 12).Value
    GTV_Ticker = ws.Cells(i, 9).Value
  End If
  
 Next i
 
 
'Print the information
 ws.Range("P2").Value = GPI_Ticker
 ws.Range("Q2").Value = GPI_Value
 ws.Range("P3").Value = GPD_Ticker
 ws.Range("Q3").Value = GPD_Value
 ws.Range("P4").Value = GTV_Ticker
 ws.Range("Q4").Value = GTV_Value
 ws.Range("Q2", "Q3").NumberFormat = "0.00%"
 
'Autofit the columns
 ws.Columns("I:Q").AutoFit

Next ws


End Sub

