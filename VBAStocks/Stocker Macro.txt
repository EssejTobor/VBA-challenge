Sub Stocker()

   
For Each ws In Worksheets
    ws.Columns(9).ClearContents
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 13).Value = "Greatest % Increase"
    ws.Cells(3, 13).Value = "Greatest % Decrease"
    ws.Cells(4, 13).Value = "Greatest Total Volume"
    ws.Cells(1, 14).Value = "Ticker"
    ws.Cells(1, 15).Value = "Value"
    Dim tablerow As Integer
    tablerow = 1
    
    Dim ticker As String
    Dim totalvalue As Double
    totalvalue = 0
    Dim yearlychange As Double
    Dim yearlyopen As Double
    Dim yearlyclose As Double
    
    
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    ws.Cells(2, 15).NumberFormat = "0.00%"
    ws.Cells(3, 15).NumberFormat = "0.00%"
    
    
For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And (ws.Cells(i, 3).Value <> 0) Then
       
        tablerow = tablerow + 1
        totalvalue = totalvalue + ws.Cells(i, 7).Value
        ws.Cells(tablerow, 12).Value = totalvalue
        totalvalue = 0
        ticker = ws.Cells(i, 1).Value
        ws.Cells(tablerow, 9).Value = ticker
        
    ElseIf ws.Cells(i, 3).Value <> 0 Then
        
        yearlychange = yearlyclose - yearlyopen
        totalvalue = totalvalue + ws.Cells(i, 7).Value
        
    End If
    
Next i

ws.Columns("K").NumberFormat = "0.00%"
Dim trow As Integer
    trow = 1
Dim percentchange As Double
Dim ychange As Double
Dim yopen As Double
Dim yclose As Double

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And (ws.Cells(i, 3).Value <> 0) Then
        trow = trow + 1
        yclose = ws.Cells(i, 6)
        ychange = yclose - yopen
        ws.Cells(trow, 10).Value = ychange
        '____"
        percentchange = ychange / yopen
        ws.Cells(trow, 11).Value = percentchange
        ychange = 0
        
        
    ElseIf (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) And (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) And (ws.Cells(i, 3).Value <> 0) Then
        
        yopen = ws.Cells(i, 3).Value
    
    Else
        
        ychange = yclose - yopen
       
    End If
    
Next i

finrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To finrow
        If ws.Cells(i, 10).Value > 0 Then
            
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 10).Value < 0 Then
            
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
Next i

Dim MinPercent As Double
    MinPercent = 0
lrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To lrow
        If ws.Cells(i, 11).Value < MinPercent Then
        
        MinPercent = ws.Cells(i, 11).Value
        ws.Cells(3, 15).Value = MinPercent
        ws.Cells(3, 14).Value = ws.Cells(i, 9).Value
        
        End If
Next i

Dim MaxPercent As Double
    MaxPercent = 0

For i = 2 To lrow
        If ws.Cells(i, 11).Value > MaxPercent Then
        
        MaxPercent = ws.Cells(i, 11).Value
        ws.Cells(2, 15).Value = MaxPercent
        ws.Cells(2, 14).Value = ws.Cells(i, 9).Value
        
        End If
Next i

Dim MaxTotal As Double
    MaxTotal = 0

For i = 2 To lrow
        If ws.Cells(i, 12).Value > MaxTotal Then
        
        MaxTotal = ws.Cells(i, 12).Value
        ws.Cells(4, 15).Value = MaxTotal
        ws.Cells(4, 14).Value = ws.Cells(i, 9).Value
        
        End If
Next i

ws.Columns("A:P").AutoFit

Next ws

    
End Sub
