Attribute VB_Name = "Module11"
Sub stockinfo()
'Variables
Dim ticker As String
Dim lastrow As Long
Dim i As Long
Dim vol As Double
Dim outputrow As Long
Dim openvaluestart As Double
Dim closevalueend As Double
Dim ychange As Double
Dim pchange As Double
Dim outputlastrow As Long


lastrow = Cells(Rows.Count, 1).End(xlUp).Row
outputlastrow = Cells(Rows.Count, 9).End(xlUp).Row
outputrow = 2
vol = 0
openvaluestart = Cells(2, 3).Value

'Listing ticker symbols and volume

For Each ws In Worksheets
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closevalueend = ws.Cells(i, 6).Value
            
            'Output Ticker Symbol in Ticker column i
            ws.Cells(outputrow, 9).Value = ws.Cells(i, 1)
            
           'Outputting Volume to Total Stock Volume Column L
            vol = vol + ws.Cells(i, 7)
            ws.Cells(outputrow, 12).Value = vol
            
            'Percent change calculation and output
           If openvaluestart = 0 Then
                pchange = 0
                Else
                pchange = (closevalueend - openvaluestart) / openvaluestart
            End If
            ws.Cells(outputrow, 11).Value = pchange
                
                'conditional formatting
                If pchange > 0 Then
                    ws.Cells(outputrow, 11).Interior.Color = 13561798
                    ws.Cells(outputrow, 11).Font.Color = 24832
                Else
                    ws.Cells(outputrow, 11).Interior.Color = 13551615
                    ws.Cells(outputrow, 11).Font.Color = 393372
                End If
            
            'Yearly Change calculation and output
            ychange = closevalueend - openvaluestart
            ws.Cells(outputrow, 10).Value = ychange
                
                'Conditional formatting
                If ychange > 0 Then
                    ws.Cells(outputrow, 10).Interior.Color = 13561798
                    ws.Cells(outputrow, 10).Font.Color = 24832
                Else
                    ws.Cells(outputrow, 10).Interior.Color = 13551615
                    ws.Cells(outputrow, 10).Font.Color = 393372
                End If
                
            openvaluestart = ws.Cells(i + 1, 6).Value
            outputrow = outputrow + 1
            vol = 0
        Else
            'Add That rows volume to running total
            vol = vol + ws.Cells(i, 7)
        End If
    Next i
        
    'greatest increase
    ws.Range("p2").Value = WorksheetFunction.Max(ws.Range("k:k"))
    ws.Range("o2").Value = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("k:k")), ws.Range("K:K"), 0), 9)
    'greatest decrease
    ws.Range("p3").Value = WorksheetFunction.Min(ws.Range("k:k"))
    ws.Range("o3").Value = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("k:k")), ws.Range("K:K"), 0), 9)
    'greatest volume
    ws.Range("p4").Value = WorksheetFunction.Max(ws.Range("l:l"))
    ws.Range("o4").Value = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("l:l")), ws.Range("l:l"), 0), 9)
        
    'Formatting
    ws.Range("a1:g1").Interior.Color = 5296274
    ws.Range("i1:l1").Interior.Color = 16764057
    ws.Range("n1:p1").Interior.Color = 16751001
    
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"
    ws.Range("n2").Value = "Greatest % Increase"
    ws.Range("n3").Value = "Greatest % Decrease"
    ws.Range("n4").Value = "Greatest Total Volume"
    
    ws.Range("C:F").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    ws.Range("J:J").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    ws.Range("G:G").NumberFormat = "#,##0_);[Red](#,##0)"
    ws.Range("L:L").NumberFormat = "#,##0_);[Red](#,##0)"
    ws.Range("p4").NumberFormat = "#,##0_);[Red](#,##0)"
    ws.Range("k:K").NumberFormat = "0.00%"
    ws.Range("p2,p3").NumberFormat = "0.00%"
    ws.Range("a:P").EntireColumn.AutoFit
    
    outputrow = 2
    openvaluestart = Cells(2, 3).Value
Next ws
    
End Sub


