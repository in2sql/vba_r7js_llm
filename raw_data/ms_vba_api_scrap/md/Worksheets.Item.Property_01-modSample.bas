Attribute VB_Name = "modSample"
Public Sub Sample()

    Let nCount = ThisWorkbook.Worksheets.Count
    Let nIndex = 1
    For nIndex = 1 To nCount
        Set ws = ThisWorkbook.Worksheets.Item(nIndex)
        MsgBox ws.Name
    Next
    
End Sub
