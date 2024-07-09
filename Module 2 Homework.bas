Attribute VB_Name = "Module1"
Sub Macro1()
'
' Macro1 Macro
'

'
    
    Dim ws As Worksheet
    
    Dim lastRowColA As Long
    
    Dim ticker As String
    
    Dim tickerCol As Long
    tickerCol = 2
        
    Dim openingAmount As Double
    
    Dim closingAmount As Double
    
    Dim quarterlyChange As Double
    
    Dim percentChange As Double
    
    Dim totalVolume As Double
        
    Dim i As Long, j As Long
    
    'For Each ws In ThisWorkbook.Worksheets
        'ws.Activate
        'MsgBox "Processing Worksheet: " & ws.Name
        
   
    
    
        'Range("I1").Select
        'ActiveCell.FormulaR1C1 = "Ticker"
        'Range("J1").Select
        'ActiveCell.FormulaR1C1 = "Quarterly Change"
        'Range("K1").Select
        'ActiveCell.FormulaR1C1 = "% Change"
        'Range("L1").Select
        'ActiveCell.FormulaR1C1 = "Total Stock Volume"
        'Columns("L:L").Select
        'Selection.NumberFormat = "#,##0"
        'Columns("K:K").Select
        'Selection.NumberFormat = "0.00%"
        'Columns("J:J").Select
        'Selection.NumberFormat = "0.00"
        'Range("O2").Select
        'ActiveCell.FormulaR1C1 = "Greatest % Increase"
        'Range("O3").Select
        'ActiveCell.FormulaR1C1 = "Greatest % Decrease"
        'Range("O4").Select
        'ActiveCell.FormulaR1C1 = "Greatest Total Volume"
        'Range("P1").Select
        'ActiveCell.FormulaR1C1 = "Ticker"
        'Range("Q1").Select
        'ActiveCell.FormulaR1C1 = "Value"
        'Range("P1:Q1").Select
        'Selection.Font.Bold = True
        'With Selection
            '.HorizontalAlignment = xlCenter
            '.VerticalAlignment = xlBottom
            '.WrapText = False
            '.Orientation = 0
            '.AddIndent = False
            '.IndentLevel = 0
            '.ShrinkToFit = False
            '.ReadingOrder = xlContext
            '.MergeCells = False
        'End With
        'Range("O2:O4").Select
        'Selection.Font.Bold = True
        'Columns("O:O").ColumnWidth = 19.64
        'Range("P2").Select

    
        
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
        'MsgBox "Processing Worksheet: " & ws.Name
        
 
    
    With ws
        
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "% Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("L:L").NumberFormat = "#,##0"
    Columns("K:K").NumberFormat = "0.00%"
    Columns("J:J").NumberFormat = "0.00"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("P1:Q1").Font.Bold = True
    
    With Range("P1:Q1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Range("O2:O4").Font.Bold = True
    Columns("O:O").ColumnWidth = 19.64
        
        
        lastRowColA = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        tickerCol = 2
        For i = 2 To lastRowColA
           
                'Find all tickers of the same name
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    openingAmount = ws.Cells(i, 3).Value
                    totalVolume = ws.Cells(i, 7).Value
                
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & tickerCol).Value = ticker
                    'ws.Cells("I", tickerCol).Value = ticker
                    closingAmount = ws.Cells(i, 6).Value
                    quarterlyChange = closingAmount - openingAmount
                    ws.Range("J" & tickerCol).Value = quarterlyChange
                    
                                       
                    'Calculate % Change and Total Volume and put in table
                    'percentChange = ws.Cells("J", tickerCol).Value / openingPrice
                    percentChange = ws.Cells(tickerCol, 10).Value / openingAmount
                    
                    'ws.Cells("K", tickerCol).Value = percentChange
                    ws.Cells(tickerCol, 11).Value = percentChange
                        
                    'totalVolume = totalVolume + ws.Cells(i, 7).Value
                        
                    'ws.Cells("L", tickerCol).Value = totalVolume
                    ws.Cells(tickerCol, 12).Value = totalVolume
                        
                    tickerCol = tickerCol + 1
                        
                    'totalVolume = Cells(i, 7).Value
            
                Else
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
            
                End If
        
            Next i
        
            'Getting info for small summary table
            Dim lastRowColK As Long
            lastRowColK = ws.Cells(ws.Rows.Count, 11).End(xlUp).row
                
            'Greatest % Increase & Decrease
            For i = 2 To lastRowColK
        
                If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2") = ws.Cells(i, 11).Value
                    ws.Range("P2") = ws.Cells(i, 9).Value
                    End If
                        
                            
                If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3") = ws.Cells(i, 11).Value
                    ws.Range("P3") = ws.Cells(i, 9).Value
                        
                End If
            
            Next i
        
        
            'Greatest Total Volume
            Dim lastRowColL As Long
            lastRowColL = ws.Cells(ws.Rows.Count, 12).End(xlUp).row
                
            For i = 2 To lastRowColL
        
                If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4") = ws.Cells(i, 12).Value
                    ws.Range("P4") = ws.Cells(i, 9).Value
                    
                End If
            
            Next i
    
        End With
    Next ws


    
End Sub


