Sub Stock_Info()

'Set variable for Stock Ticker
Dim Stock_Ticker As String

'Set variables for Opening Price
Dim Starting As Long

'Set variable for Closing Price
Dim Closing As Long

'Set variable for Total Stock Volume
Dim Total As Long

'Set variable for Percentage Change
Dim Percentage_Change As Long

'Set Summary Table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'Set Titles
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


'Loop through all stock information
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
    'Check for changes in ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set Stock_Ticker
       Stock_Ticker = Cells(i, 1).Value
    
   'Print
         Range("I" & Summary_Table_Row).Value = Stock_Ticker
         
    'Add Row to table
         Summary_Table_Row = Summary_Table_Row + 1
     
   'Calculate total volume
         Total = Total_Stock_Volume + Cells(i, 6).Value
         
     'Print
          Range("L" & 2).Value = Total_Stock_Value
    
    'Calculate and Store the yearly change
    
          Yearly_Change = Closing_Price - Opening_Price
          Range("J" & 2).Value = Yearly_Change
    
   
          
    'Reset total Stock Volume
          Total_Stock_Volume = 0
          
  
    End If
 Next i

End Sub
