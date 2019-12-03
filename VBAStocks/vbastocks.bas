Attribute VB_Name = "Module1"
Sub SummarizeStocks()
    For Each ws In Worksheets
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        outrow = 2
        yearstartval = ws.Cells(2, 3).Value
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        totalvol = 0
        For i = 2 To lastrow
            totalvol = totalvol + ws.Cells(i, 7).Value
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                yearendval = ws.Cells(i, 6).Value
                yearchg = yearendval - yearstartval
                If yearstartval > 0 Then
                    yearpchg = yearchg / yearstartval
                Else
                    yearpchg = Null
                End If
                
                ws.Cells(outrow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(outrow, 10).Value = yearchg
                ws.Cells(outrow, 11).Value = yearpchg
                ws.Cells(outrow, 11).NumberFormat = "0.00%"
                ws.Cells(outrow, 12).Value = totalvol
                
                outrow = outrow + 1
                yearstartval = ws.Cells(i + 1, 3).Value
                totalvol = 0
            End If
        Next i
        
        greatestpctinc = 0
        greatestpctdec = 0
        greatesttotvol = 0
        
        For j = 2 To (outrow - 1)
            If ws.Range("J" & j) < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3
            ElseIf ws.Range("J" & j) > 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 4
            Else
                ws.Range("J" & j).Interior.ColorIndex = 15
            End If
            
            
            If ws.Range("K" & j).Value > greatestpctinc Then
                greatestpctinc = ws.Range("K" & j).Value
                greatestpctinctick = ws.Range("I" & j)
            End If
            If ws.Range("K" & j).Value < greatestpctdec Then
                greatestpctdec = ws.Range("K" & j).Value
                greatestpctdectick = ws.Range("I" & j)
            End If
            If ws.Range("L" & j).Value > greatesttotvol Then
                greatesttotvol = ws.Range("L" & j).Value
                greatesttotvoltick = ws.Range("I" & j)
            End If
            
        Next j
        
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatset Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ws.Range("O2").Value = greatestpctinctick
        ws.Range("P2").Value = greatestpctinc
        ws.Range("O3").Value = greatestpctdectick
        ws.Range("P3").Value = greatestpctdec
        ws.Range("O4").Value = greatesttotvoltick
        ws.Range("P4").Value = greatesttotvol
        
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
    Next
End Sub
