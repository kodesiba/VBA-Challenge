Attribute VB_Name = "Module1"
Sub SummarizeStocks()
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    outrow = 2
    yearstartval = Cells(2, 3).Value
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    totalvol = 0
    For i = 2 To lastrow
        totalvol = totalvol + Cells(i, 7).Value
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            yearendval = Cells(i, 6).Value
            yearchg = yearendval - yearstartval
            If yearstartval > 0 Then
                yearpchg = yearchg / yearstartval
            Else
                yearpchg = Null
            End If
            
            Cells(outrow, 9).Value = Cells(i, 1).Value
            Cells(outrow, 10).Value = yearchg
            Cells(outrow, 11).Value = yearpchg
            Cells(outrow, 11).NumberFormat = "0.00%"
            Cells(outrow, 12).Value = totalvol
            
            outrow = outrow + 1
            yearstartval = Cells(i + 1, 3).Value
            totalvol = 0
        End If
    Next i
    
    greatestpctinc = 0
    greatestpctdec = 0
    greatesttotvol = 0
    
    For j = 2 To (outrow - 1)
        If Range("J" & j) < 0 Then
            Range("J" & j).Interior.ColorIndex = 3
        ElseIf Range("J" & j) > 0 Then
            Range("J" & j).Interior.ColorIndex = 4
        Else
            Range("J" & j).Interior.ColorIndex = 15
        End If
        
        
        If Range("K" & j).Value > greatestpctinc Then
            greatestpctinc = Range("K" & j).Value
            greatestpctinctick = Range("I" & j)
        End If
        If Range("K" & j).Value < greatestpctdec Then
            greatestpctdec = Range("K" & j).Value
            greatestpctdectick = Range("I" & j)
        End If
        If Range("L" & j).Value > greatesttotvol Then
            greatesttotvol = Range("L" & j).Value
            greatesttotvoltick = Range("I" & j)
        End If
        
    Next j
    
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatset Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    Range("O2").Value = greatestpctinctick
    Range("P2").Value = greatestpctinc
    Range("O3").Value = greatestpctdectick
    Range("P3").Value = greatestpctdec
    Range("O4").Value = greatesttotvoltick
    Range("P4").Value = greatesttotvol
    
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
End Sub
