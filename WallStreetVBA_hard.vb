' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'            UNIT 2 - VBA Assignment - Hard 
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #

Sub BetaWallStreet()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    'Changed the headers because that was bugging me
    ws.Range("A1").Value = "ticker"
    ws.Range("B1").Value = "date"
    ws.Range("C1").Value = "open"
    ws.Range("D1").Value = "high"
    ws.Range("E1").Value = "low"
    ws.Range("F1").Value = "close"
    ws.Range("G1").Value = "vol"

    'Headers for the new data columns we're making
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "% Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Variables for the looping
    Dim ticker As String    'name of the ticker we're working on
    Dim var As Double       '+1 for each new ticker
    Dim sum As LongLong     'adds up each tickers total volume
    Dim LastRow As LongLong 'determines # of last row
    Dim initial As Double   'records year open value
    Dim final As Double     'records year close value

    'Set initial values for variables
    ticker = ws.Range("A2").Value
    var = 2
    sum = ws.Range("G2").Value
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("I2").Value = ticker
    initial = ws.Range("C2").Value


    For i = 3 To LastRow + 1
        ticker = ws.Cells(i, 1).Value
        
        If ticker = ws.Cells(i - 1, 1).Value Then
            sum = sum + ws.Cells(i, 7).Value
        Else
            ws.Cells(var, 12).Value = sum
            final = ws.Cells(i - 1, 6).Value
            ws.Cells(var, 10).Value = final - initial
            ws.Cells(var, 10).NumberFormat = "0.00"
            
            If ws.Cells(var, 10).Value > 0 Then
                ws.Cells(var, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(var, 10).Value < 0 Then
                ws.Cells(var, 10).Interior.ColorIndex = 3
            End If
            
            If initial <> 0 Then
                ws.Cells(var, 11).Value = ws.Cells(var, 10).Value / initial
                ws.Cells(var, 11).NumberFormat = "0.000%"
            Else
                ws.Cells(var, 11).Value = 0
            End If
            
            var = var + 1
            initial = ws.Cells(i, 3)
            ws.Cells(var, 9).Value = ticker
            sum = ws.Cells(i, 7).Value
        End If
        
    Next i

    
    'Some new variables to determine greatest % increase/decrease & tot vol
    Dim greatest_inc As Double
    Dim greatest_dec As Double
    Dim greatest_vol As LongLong

    greatest_inc = 0
    greatest_dec = 0
    greatest_vol = 0


    'Loop through and record the values we want and 
    'their corresponding tickers
    For j = 2 To var
        If ws.Cells(j, 11).Value > greatest_inc Then    
            greatest_inc = ws.Cells(j, 11).Value
            ws.Range("O2") = ws.Cells(j, 9)
        End If
        If ws.Cells(j, 11).Value < greatest_dec Then
            greatest_dec = ws.Cells(j, 11).Value
            ws.Range("O3") = ws.Cells(j, 9)
        End If
        If ws.Cells(j, 12).Value > greatest_vol Then
            greatest_vol = ws.Cells(j, 12).Value
            ws.Range("O4") = ws.Cells(j, 9)
        End If
    Next j

    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest total volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("P2").Value = greatest_inc
    ws.Range("P3").Value = greatest_dec
    ws.Range("P2:P3").NumberFormat = "0.000%"
    ws.Range("P4").Value = greatest_vol

Next ws

End Sub
