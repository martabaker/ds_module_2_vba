Sub reset():
    'Meant to reset both the contents and color the columns where the ticker summary columns will be
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ws.Range("I:P").Value = ""
        ws.Range("I:L").Interior.ColorIndex = 2
        ws.Range("I:P").ColumnWidth = 8.43

    Next ws

End Sub