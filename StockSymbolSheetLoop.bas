Attribute VB_Name = "Module1"
Sub StockSymbolSheetLoop()

Dim ws As Worksheet
         
For Each ws In ThisWorkbook.Sheets
    
''''''''''''''''''''''''''''''''''''''''''''''''''
           
'Write in headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Set an initial variable for holding the stock symbol
  Dim symbol As String

  ' Set an initial variable for holding the total per stock symbol
  Dim Total_volume As Double
  Total_volume = 0
  
  Dim percent_change As Double

  ' Keep track of the location for each stock symbol brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Find last row
  lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Loop through all ticker symbols purchases
  For i = 2 To lastRow

    ' Check to see if this is the last record with the current stock symbol
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Set the stock symbol name
        symbol = ws.Cells(i, 1).Value
        
        'Add to the Total Volume
        Total_volume = Total_volume + ws.Cells(i, 7).Value
        
        'Print the symbol in summary table
        ws.Range("I" & Summary_Table_Row).Value = symbol

      
        'Print the Total Volume to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Total_volume
              
        'Zero out total volume at the very last row
        Total_volume = 0
        
        'locate close number
        Close_Num = ws.Cells(i, 6).Value
        
        
        'Calculate Yearly Change
        
        Yearly_Change = (Close_Num - Open_Num)
        
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
            
        'Calculate Percent Change
        If Open_Num <> 0 Then
        percent_change = (Yearly_Change / Open_Num)
                     
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
          
         Else
         
         ws.Range("K" & Summary_Table_Row).Value = "0"
         
         End If
        
        ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'Check to see if this is the first record of this symbol type
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'locate open number
        Open_Num = ws.Cells(i, 3).Value
            
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Check if the cell immediately following a row is the same stock symbol
    Else
     
      
      ' Add to the Total Volume
      Total_volume = Total_volume + ws.Cells(i, 7).Value
      
     
         
    End If
 ''''''''''''''''' '''''''''''''''''''''''''''''''''''
    

''''''''''''loop through active sheet''''''''''''''''''
  Next i
           
'Find last row
  Summary_lastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
  
  For j = 2 To Summary_lastRow
  
  'Conditional formatting the Yearly Change
    If ws.Cells(j, 10).Value >= "0" Then
           
    'Formatting positive change with green interior
    ws.Cells(j, 10).Interior.ColorIndex = 43
           
    Else
    'Formatting negative change with red interior
    ws.Cells(j, 10).Interior.ColorIndex = 3

    End If
    
    Next j
           
'''''''''''''loop through worksheet''''''''''''''''''''
Next ws

End Sub

