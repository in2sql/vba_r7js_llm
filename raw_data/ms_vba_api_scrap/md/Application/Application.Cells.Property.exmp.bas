Sub ChangeInsertRows()
    Application.ScreenUpdating = False
    Dim xRow As Long
    
    For xRow = Application.Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
        If Application.Cells(xRow, 1).Value <> Application.Cells(xRow - 1, 1).Value Then Rows(xRow).Resize(1).Insert
    Next xRow
    
    Application.ScreenUpdating = True
End Sub