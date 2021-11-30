Sub TickerandTotals()

' Set variables
Dim ticker As String
Dim total As Double
Dim ws As Worksheet

'Run for all worksheets
For Each ws In Worksheets
    
    'set variables for the each worksheet
    total = 0
    j = 0


'Set up the results chart
ws.Range("I1").Value = "Ticker"
ws.Range("L1").Value = "Total Stock Volume"



'Loop through all stock information
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'If the value of a cell in column A is not equal to the value below it
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Print the ticker symbol in that row
        ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
        
        'Print the total value for that ticker
        ws.Range("l" & 2 + j).Value = total
        
        'Next reset the starting total to zero
        total = 0
        
        'Go to the next row and start again
        j = j + 1
        
    'If the stock ticker is the same as the one below it, add the volume amount on that row to the total.
    Else
        total = total + ws.Cells(i, 7).Value
    End If
    
Next i

Next ws

End Sub

Sub Changes()
'set up variables
Dim i As Long
Dim annual_chg As Single
Dim j As Integer
Dim start As Long
Dim pc_chg As Single
Dim total As Double
Dim ws As Worksheet

'Run for all worksheets
For Each ws In Worksheets

'set remaining titles
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"

'Set initial values
j = 0
annual_chg = 0
start = 2


'Loop through all stock information
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
      'If the value of a cell in column A is not equal to the value below it
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Store total in variables
        total = total + ws.Cells(i, 7).Value
        
        'Handle the zero balances
         If total = 0 Then
            'Print the results
            ws.Range("J" & 2 + j).Value = "%" & 0
            ws.Range("K" & 2 + j).Value = 0
          Else
        
            'Find the first nonzero starting value for the stock
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                        Exit For
                    End If
                Next find_value
            End If
            
            'Calculate the yearly change
            annual_chg = (ws.Cells(i, 6) - ws.Cells(start, 3))
            pc_chg = Round((annual_chg / ws.Cells(start, 3) * 100), 2)
        
        'Start at the next unique stock ticker
        start = i + 1
        
        'Print the changes
        ws.Range("J" & 2 + j).Value = Round(annual_chg, 2)
        ws.Range("K" & 2 + j).Value = "%" & pc_chg
        
        
        'Make the positives green and the negatives red
        Select Case annual_chg
            Case Is > 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
                ws.Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
            
    End If
    
    'Reset variables for next stock ticker
    total = 0
    annual_chg = 0
    j = j + 1
    
    'If ticker is the same add to the total volume
    Else
        total = total + ws.Cells(i, 7).Value
  
    End If
    
Next i

Next ws
    
    
End Sub
