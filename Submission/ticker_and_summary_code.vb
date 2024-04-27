Sub ticker():

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate 'for all worksheets
        
        'Add column headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Formatting for the final columns
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("J:J").ColumnWidth = 16
        ws.Range("K:K").ColumnWidth = 15
        ws.Range("L:L").ColumnWidth = 18
        
        'Add variables
        Dim i As Long 'row number
        Dim vol As Double 'contents of column G
        Dim totalvol As Double 'goes in column L
        Dim ticker As String 'goes in column I
        Dim j As Integer 'This will be the row where values are placed
        Dim FinalRow As Long
        Dim lastRow As Long
        Dim startRow As Long
        Dim endRow As Long
        
        FinalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Final Row in the column A
        
        
        totalvol = 0 'starting this value at because at the beginning of the loop, there has been no stocks traded
        startRow = 2
        j = 2
        
        For i = 2 To FinalRow
            vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                totalvol = (totalvol + vol)
                
                endRow = i
                Dim firstOpen As Double 'Used for quarterchange and percentchange
                Dim finalClose As Double 'Used for quarterchange and percentchange
                Dim percentChange As Double 'goes in column K
                Dim quarterChange As Double 'goes in column J
                
                firstOpen = ws.Cells(startRow, 3).Value ' Assuming oldPrice is in column c
                finalClose = ws.Cells(endRow, 6).Value ' Assuming newPrice is in column F
                
                ' Calculate the percent change
                percentChange = ((finalClose - firstOpen) / firstOpen)
                
                ' Calculate the quarter change
                quarterChange = (finalClose - firstOpen)
                
                'Need to add the values where the go in the table
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = quarterChange
                ws.Cells(j, 11).Value = percentChange
                ws.Cells(j, 12).Value = totalvol
                
                'conditional formatting for the percentChange and quarterChange
                If (percentChange > 0) Then
                    ws.Cells(j, 11).Interior.ColorIndex = 4
                    
                ElseIf (percentChange < 0) Then
                    ws.Cells(j, 11).Interior.ColorIndex = 3
                
                End If
                
                If (quarterChange > 0) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    
                ElseIf (quarterChange < 0) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                
                End If
                
                
                'reset
                totalvol = 0
                j = j + 1
                quarterChange = 0
                percentChange = 0
                startRow = i + 1
                
                
                
            Else
                'ONLY need to add total
                totalvol = totalvol + vol
            
            End If
            
        Next i
        
    'The following provides the summary information
        'Add Cell names
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Value"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'Formatting
        ws.Range("P2:P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "0"
        ws.Range("N:N").ColumnWidth = 20
        ws.Range("P:P").ColumnWidth = 20
        
        
        Dim tickermaxp As String 'ticker for the max % increase
        Dim tickerminp As String 'ticker for the max % decrease
        Dim tickermaxv As String 'ticker for the greatest total volume
        Dim maxvol As Double 'Greatest Total Value
        Dim maxp As Double 'Greatest % Increase value
        Dim minp As Double 'Greatest % Decrease value
        Dim k As Long 'To find ticker for minp
        Dim l As Long 'To find ticker for maxvol
                    
        maxp = WorksheetFunction.Max(Range("K:K"))
        minp = WorksheetFunction.Min(Range("K:K"))
        maxvol = WorksheetFunction.Max(Range("L:L"))
        
        FinalRow = ws.Cells(Rows.Count, 9).End(xlUp).Row 'Final Row in column I
        
        j = 2
        
        For i = 2 To FinalRow
        
            If (ws.Cells(i, 11).Value = maxp) Then
                ws.Cells(j, 15).Value = Cells(i, 9).Value
                ws.Cells(j, 16).Value = maxp
                
            Else
            
            End If
        
        Next i
        
        j = j + 1
        
        For k = 2 To FinalRow
            If (ws.Cells(k, 11).Value = minp) Then
                ws.Cells(j, 15).Value = Cells(k, 9).Value
                ws.Cells(j, 16).Value = minp
    
            Else
            
            End If
            
        Next k
        
        j = j + 1
        
        For l = 2 To FinalRow
            If (ws.Cells(l, 12).Value = maxvol) Then
                ws.Cells(j, 15).Value = Cells(l, 9).Value
                ws.Cells(j, 16).Value = maxvol
            
            Else
            
            End If
            
        
        Next l

    Next ws

End Sub