Attribute VB_Name = "Module1"
Sub Ticker()
Dim Ticker As String
Dim opening, closing, year_change, totalstockvol, perc_change As Double
Dim starting As Integer
Dim ws As Worksheet

For Each ws In Worksheets

'Add columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'assign starting integer
starting = 2
Count = 1
totalstockvol = 0

'set the last row
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastrow
   'it Ticker is not the same as the row before
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Get Ticker in order
            Ticker = ws.Cells(i, 1).Value
    
    'get the i to count up by one each time
            Count = Count + 1
    
    'get the value from first day open and last day close
    opening = ws.Cells(Count, 3).Value
    closing = ws.Cells(i, 6).Value
    
    'sum the total stock volume
    For j = Count To i
        totalstockvol = totalstockvol + ws.Cells(j, 7).Value
    Next j
    
    'open the data
    If opening = 0 Then
        perc_change = closing
Else
    year_change = closing - opening
    perc_change = year_change / opening
End If
    'print in sum table
    ws.Cells(starting, 9).Value = Ticker
    ws.Cells(starting, 10).Value = year_change
    ws.Cells(starting, 11).Value = perc_change
    ws.Cells(starting, 12).Value = totalstockvol
    
    'go to the next row
    starting = starting + 1
    
    'reset the values
    totalstockvol = 0
    year_change = 0
    perc_change = 0
    
    'reset the count
    Count = i
End If
Next i

'next summary table
'greatest increase is column K 11
'greatest decrease is column K 11
'greatest total volume is column L12
 ' Assign names for summary table 2
    
    ws.Range("O1").Value = "Greatest % Increase"
    ws.Range("P1").Value = "Greatest % Decrease"
    ws.Range("Q1").Value = "Greatest Total Volume"
    ws.Range("N2").Value = "Ticker"
    ws.Range("N3").Value = "Value"
    
   



'start with column k
    klastrow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
        For k = 3 To klastrow
        'check the previous value
        previous_k = k - 1
        
        'current k value for %
        current = ws.Cells(k, 11).Value
        
        'previous k value for %
        Previous = ws.Cells(previous_k, 11).Value
        
'set baseline value for asks: greatest increase, greatest decrease, greatest total stock volume
perc_increase = 0
perc_decrease = 0
great_total_vol = 0

'Find the % increase

If perc_increase > current And perc_increase > previous_k Then
perc_increase = perc_increase

great_increase_name = ws.Cells(k, 9).Value

ElseIf current > perc_increase And current > previous_k Then
perc_increase = current
great_increase_name = ws.Cells(k, 9).Value

ElseIf previous_k > current And previous_k > perc_increase Then
perc_increase = previous_k
great_increase_name = ws.Cells(k, 9).Value

End If

'Find the greatest decrease
If perc_decrease < current And perc_decrease < previous_k Then
perc_decrease = perc_decrease

great_decrease_name = ws.Cells(k, 9).Value

ElseIf current < perc_decrease And current < previous_k Then
perc_decrease = current
great_decrease_name = ws.Cells(k, 9).Value

ElseIf previous_k < current And previous_k < perc_decrease Then
perc_decrease = previous_k
great_decrease_name = ws.Cells(k, 9).Value

End If
Next k

'use column l for greatest total volume
       llastrow = ws.Cells(Rows.Count, "L").End(xlUp).Row
        
        For l = 3 To llastrow
        'check the previous value
        previous_l = l - 1
        
        'current l value for volume
        current_l = ws.Cells(l, 12).Value
        
        'previous l volume for volume
        volume = ws.Cells(previous_l, 12).Value
        
'find greatest total volume
If great_total_vol > current_l And great_total_vol > previous_l Then
great_total_vol = great_total_vol
great_total_vol_name = ws.Cells(l, 9).Value

ElseIf current_l > great_total_vol And current_l > previous_l Then
great_total_vol = current_l
great_total_vol_name = ws.Cells(l, 9).Value

ElseIf previous_l > current_l And previous_l > great_total_vol Then
great_total_vol = previous_l
great_total_vol_name = ws.Cells(l, 9).Value

End If
Next l

 'Input values for summary table 2
    ws.Range("O2").Value = great_increase_name
    ws.Range("P2").Value = great_decrease_name
    ws.Range("Q2").Value = great_total_vol_name
    ws.Range("O3").Value = perc_increase
    ws.Range("P3").Value = perc_decrease
    ws.Range("Q3").Value = great_total_vol
'Conditional formatting columns colors for yearly change

    jlastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
        For j = 2 To jlastrow
            
            'if conditional formatting
            If ws.Cells(j, 10) > 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
Next ws
End Sub
