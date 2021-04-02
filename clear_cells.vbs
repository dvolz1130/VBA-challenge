Sub clear_cells()
    For Each ws In Worksheets
        ws.Columns(9).ClearContents
        ws.Columns(10).ClearContents
        ws.Columns(10).Interior.Color = xlNone
        ws.Columns(11).ClearContents
        ws.Columns(11).NumberFormat = "General"
        ws.Columns(12).ClearContents
        
        ' If I forget to clear before re-running
        ws.Columns(14).ClearContents
        ws.Columns(15).ClearContents
        ws.Columns(15).Interior.Color = xlNone
        ws.Columns(16).ClearContents
        ws.Columns(16).NumberFormat = "General"
        ws.Columns(17).ClearContents
    Next ws
End Sub