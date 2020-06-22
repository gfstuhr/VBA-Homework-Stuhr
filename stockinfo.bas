Attribute VB_Name = "Module1"
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


lastrow = Cells(Rows.Count, 1).End(xlUp).row
outputlastrow = Cells(Rows.Count, 9).End(xlUp).row
outputrow = 2
vol = 0
openvaluestart = Cells(2, 3).Value

'Listing ticker symbols and volume
For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        closevalueend = Cells(i, 6).Value
        
        'Output Ticker Symbol in Ticker column i
        Cells(outputrow, 9).Value = Cells(i, 1)
        
       'Outputting Volume to Total Stock Volume Column L
        vol = vol + Cells(i, 7)
        Cells(outputrow, 12).Value = vol
        
        'Percent change calculation and output
       If openvaluestart = 0 Then
            pchange = 0
            Else
            pchange = (closevalueend - openvaluestart) / openvaluestart
        End If
        Cells(outputrow, 11).Value = pchange
            
            'conditional formatting
            If pchange > 0 Then
                Cells(outputrow, 11).Interior.Color = 13561798
                Cells(outputrow, 11).Font.Color = 24832
            Else
                Cells(outputrow, 11).Interior.Color = 13551615
                Cells(outputrow, 11).Font.Color = 393372
            End If
        
        'Yearly Change calculation and output
        ychange = closevalueend - openvaluestart
        Cells(outputrow, 10).Value = ychange
            
            'Conditional formatting
            If ychange > 0 Then
                Cells(outputrow, 10).Interior.Color = 13561798
                Cells(outputrow, 10).Font.Color = 24832
            Else
                Cells(outputrow, 10).Interior.Color = 13551615
                Cells(outputrow, 10).Font.Color = 393372
            End If
            
        openvaluestart = Cells(i + 1, 6).Value
        outputrow = outputrow + 1
    Else
        'Add That rows volume to running total
        vol = vol + Cells(i, 7)
    End If
Next i
    
'greatest increase
Range("p2").Value = WorksheetFunction.Max(Range("k:k"))
Range("o2").Value = Cells(WorksheetFunction.Match(WorksheetFunction.Max(Range("k:k")), Range("K:K"), 0), 9)
'greatest decrease
Range("p3").Value = WorksheetFunction.Min(Range("k:k"))
Range("o3").Value = Cells(WorksheetFunction.Match(WorksheetFunction.Min(Range("k:k")), Range("K:K"), 0), 9)
'greatest volume
Range("p4").Value = WorksheetFunction.Max(Range("l:l"))
Range("o4").Value = Cells(WorksheetFunction.Match(WorksheetFunction.Max(Range("l:l")), Range("l:l"), 0), 9)
    
'Formatting
Range("a1:g1").Interior.Color = 5296274
Range("i1:l1").Interior.Color = 16764057
Range("n1:p1").Interior.Color = 16751001

Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("l1").Value = "Total Stock Volume"
Range("o1").Value = "Ticker"
Range("p1").Value = "Value"
Range("n2").Value = "Greatest % Increase"
Range("n3").Value = "Greatest % Decrease"
Range("n4").Value = "Greatest Total Volume"

Range("C:F").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("J:J").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("G:G").NumberFormat = "#,##0_);[Red](#,##0)"
Range("L:L").NumberFormat = "#,##0_);[Red](#,##0)"
Range("p4").NumberFormat = "#,##0_);[Red](#,##0)"
Range("k:K").NumberFormat = "0.00%"
Range("p2,p3").NumberFormat = "0.00%"
Range("a:P").EntireColumn.AutoFit


    
End Sub
