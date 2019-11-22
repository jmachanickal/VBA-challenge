Attribute VB_Name = "Module1"
Sub VBAStocks()

Dim i, roww, j, h As Integer
Dim opening, closing, jump, decline As Double
Dim total, vol, largeTotal, stub As Long
Dim ticker, tickInc, tickDec, tickTotal As String
Dim ws As Worksheet

'Go through each worksheet
For Each ws In ActiveWorkbook.Worksheets
    
    'Set the labels'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Set the default Ticker and Opening stock market price variables from the first row of raw data
    ticker = ws.Cells(2, 1).Value
    opening = ws.Cells(2, 3).Value
    total = ws.Cells(2, 7).Value
    
    'Set these varibles to default values
    vol = 0
    roww = 2
    
    'Set NumRows for the loop to stop at an empty cell
    NumRows = ws.Range("A3", ws.Range("A2").End(xlDown)).Rows.Count
    'Select cell A3
    Range("A3").Select
    
    'Use for loop to go through each row starting at the second row up til an empty cell
    For i = 3 To NumRows
        
        'If the new ticker equals the previous ticker, add to the total
        If ws.Cells(i, 1).Value = ticker Then
            closing = ws.Cells(i, 6).Value
            vol = ws.Cells(i, 7).Value
            total = total + vol
        
        'If the new ticker is not equal to the previous ticker, input the variables and start variables allover
        ElseIf ws.Cells(i, 1).Value <> ticker Then
            'Fill in columns I-L
            ws.Cells(roww, 9).Value = ticker
            ws.Cells(roww, 10).Value = closing - opening
            If opening = 0 Then
                ws.Cells(roww, 11).Value = 0
            Else
                ws.Cells(roww, 11).Value = Round(((closing - opening) / opening) * 100, 2)
            End If
            ws.Cells(roww, 12).Value = total
            
            'Fill in the Yearly Change interior color based on +/- value
            If ws.Cells(roww, 10).Value > 0 Then
                ws.Cells(roww, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(roww, 10).Value < 0 Then
                ws.Cells(roww, 10).Interior.ColorIndex = 3
            End If
            
            ticker = ws.Cells(i, 1).Value
            opening = ws.Cells(i, 3).Value
            total = ws.Cells(i, 7).Value
            roww = roww + 1
        End If
    Next i
Next

For Each ws In ActiveWorkbook.Worksheets
    
    'Set the labels
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    'Set both the default percentange increase and decrease values
    If ws.Range("K2").Value > ws.Range("K3").Value Then
        jump = ws.Range("K2").Value
        tickInc = ws.Range("I2").Value
        decline = ws.Range("K3").Value
        tickDec = ws.Range("I3").Value
    ElseIf ws.Range("K2").Value < ws.Range("K3").Value Then
        decline = ws.Range("K2").Value
        tickDec = ws.Range("I2").Value
        jump = ws.Range("K3").Value
        tickInc = ws.Range("I3").Value
    Else
        jump = ws.Range("K2").Value
        tickInc = ws.Range("I2").Value
        decline = ws.Range("K2").Value
        tickDec = ws.Range("I2").Value
    End If

    'Set the default value for the greatest total volume
    tickTotal = ws.Range("I2").Value
    largeTotal = ws.Range("L2").Value
    
    'Set NumRows for the loop to stop at an empty cell
    NumRows = ws.Range("K4", ws.Range("K2").End(xlDown)).Rows.Count
    'Select cell K4
    Range("K4").Select
    
    'Check if the percetange change is higher or lower than the default value(s)
    For j = 4 To NumRows
        If ws.Cells(j, 11).Value > jump Then
            jump = ws.Cells(j, 11).Value
            tickInc = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 11).Value < decline Then
            decline = ws.Cells(j, 11).Value
            tickDec = ws.Cells(j, 9).Value
        End If
    Next j

    'Set NumRows for the loop to stop at an empty cell
    NumRows2 = ws.Range("L3", ws.Range("L3").End(xlDown)).Rows.Count
    'Select cell L3
    Range("L3").Select
    
    For h = 3 To NumRows2
        If ws.Cells(h, 12).Value > largeTotal Then
            largeTotal = ws.Cells(h, 12).Value
            tickTotal = ws.Cells(h, 9).Value
        End If
    Next h

    'Insert final values into destinated spots
    ws.Range("O2").Value = tickInc
    ws.Range("O3").Value = tickDec
    ws.Range("P2").Value = jump
    ws.Range("P3").Value = decline
    ws.Range("O4").Value = tickTotal
    ws.Range("P4").Value = largeTotal
Next

End Sub
