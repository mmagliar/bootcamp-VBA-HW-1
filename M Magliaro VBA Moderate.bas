Attribute VB_Name = "Module1"


Sub stock()
Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Dim stock_close As Double
    Dim stock_change As Double
    Dim stock_percent As Double
    Dim Summary_Table_Row As Integer
    Dim stocksym As String
    Dim vol_total As Double
    Dim lastrow As Long
'--------------------------------------------------------------------------
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
'---------------------------------------------
  ' Set an initial variable for holding the total volume per stock symbol
    vol_total = 0
'-------------------------------------
'sort columns
lastrow = Cells(Rows.Count, 2).End(xlUp).Row
Range("A2:G" & lastrow).Sort key1:=Range("a2:B" & lastrow), _
   order1:=xlAscending, Header:=xlNo
'-------------------------------------
'Keep track of the location for each stock symbol
  
Summary_Table_Row = 2
'-------------------------------------
'add column headers for summary table
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"
Cells(1, 14).Value = "Stock Open Price"
Cells(1, 15).Value = "Stock Close Price"
'-------------------------------------
' Loop through all stocks
  For I = 2 To lastrow
  
    'For I = 2 To 10000
    
'Loop to get the stock open
 If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
        stock_open = Cells(I, 3).Value
        Range("N" & Summary_Table_Row) = stock_open

' Check if we are still within the same stock symbol
    ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

' Set the Stock Symbol
        stocksym = Cells(I, 1).Value
      
' Add to the Stock  Vol  Total
        vol_total = vol_total + Cells(I, 7).Value
   
' Print the Stock Ticker  in the Summary Table
      
        Range("J" & Summary_Table_Row).Value = stocksym

' Print the Stock Vol ammount to the Summary Table
      
        Range("M" & Summary_Table_Row).Value = vol_total
 
 ' Get stock close on last line
      
        stock_close = Cells(I, 6).Value
        
        Range("o" & Summary_Table_Row) = stock_close
        '-------
        
        stock_change = (stock_close - stock_open)
        Range("k" & Summary_Table_Row) = stock_change
   
'------------------------------------------------------
'Check for divide by zero

        If (stock_change) <> 0 Then
            stock_percent = ((stock_open / stock_change) / 100)
            Range("l" & Summary_Table_Row) = stock_percent
        Else
                stock_percent = 0
' Add one to the summary table row
        End If
       
    Summary_Table_Row = Summary_Table_Row + 1
        
' Reset the Stock Vol Total
       vol_total = 0
Else
  vol_total = vol_total + Cells(I, 7).Value
  
    End If
          
    Next I
'------------------------------------------------------
 

'------------------------------------------------------
' Set the Formatting of the box
  Dim lastrowsum As Long
  lastrowsum = Cells(Rows.Count, 10).End(xlUp).Row
  
 ' Loop through summary table
  For S = 2 To lastrowsum
  
  '  For S = 2 To 10000
    
  
  If Cells(S, 11).Value > 0 Then
        Cells(S, 11).Interior.ColorIndex = 4
     ElseIf Cells(S, 11).Value < 0 Then
          Cells(S, 11).Interior.ColorIndex = 3
     Else: Cells(S, 11).Interior.ColorIndex = 2
     End If
     
     Next S
  'Range("J1:M" & lastrowsum).Interior.ColorIndex = 8
  Range("J1:P1").Font.Bold = True
  Range("M1:M" & lastrowsum).NumberFormat = "#,##0"
  Range("L1:l" & lastrowsum).NumberFormat = "0.00%"
  Range("K1:k" & lastrowsum).NumberFormat = "#,##0.##"
  
      
'---------------------------------------------------------

ws.Cells(1, 1) = 1 'this sets cell A1 of each sheet to "1"
Next
starting_ws.Activate 'activate the worksheet that was originally active
For Each ws2 In ThisWorkbook.Worksheets
        ws2.Columns.AutoFit
Next
End Sub


