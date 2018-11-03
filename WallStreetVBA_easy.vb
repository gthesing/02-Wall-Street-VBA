
'Wall Street VBA Project, easy difficulty

Sub WallStreetEasy()


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
    ws.Range("J1").Value = "Total Stock Volume"

    'Variables for the looping
    Dim ticker As String    'name of the ticker we're working on
    Dim var As Double       '+1 for each new ticker
    Dim sum As LongLong     'adds up each tickers total stock volume
    Dim LastRow As LongLong 'determines # of last row

    'Set initial values for variables
    ticker = ws.Range("A2").Value
    var = 2
    sum = ws.Range("G2").Value
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("I2").Value = ticker       'first ticker name for created column
    
    For i = 3 To LastRow                            'start at 3 because we're checking against the previous row
        ticker = ws.Cells(i, 1).Value               'ticker is renamed each row
        If ticker = ws.Cells(i - 1, 1).Value Then   'check ticker name against previous row
            sum = sum + ws.Cells(i, 7).Value        'if ticker name remains the same, add vol to total stock volume sum
        Else                                        'if ticker name is different....
            ws.Cells(var, 10).Value = sum           'record the sum of the previous ticker in new column (J)
            var = var + 1                           'next row in created column
            ws.Cells(var, 9).Value = ticker         'record current ticker name in new column (I)
            sum = ws.Cells(i, 7).Value              'restart total stock vol sumation for new ticker
        End If
    Next i
    
    ws.Cells(var, 10).Value = sum       'gets the total stock vol sum for the last ticker
    
Next ws


End Sub

