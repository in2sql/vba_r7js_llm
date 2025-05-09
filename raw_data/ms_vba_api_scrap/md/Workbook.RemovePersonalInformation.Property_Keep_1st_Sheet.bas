Attribute VB_Name = "Keep_1st_Sheet"
Sub Keep_1st_Sheet()
    Dim Ws As Worksheet
    ActiveWorkbook.RemovePersonalInformation = False
    Application.DisplayAlerts = False
    Worksheets(1).Activate
    WbName = Application.ActiveWorkbook.Name
    If ActiveWorkbook.Worksheets.Count > 1 Then
        For Each Ws In ActiveWorkbook.Worksheets
            If Ws.Name <> ActiveWorkbook.ActiveSheet.Name Then Ws.Delete
        Next Ws
    End If
    If Range("A2").Value = "Work_Day" Then
        Range("A2", Range("A2").End(xlDown).End(xlToRight)).Cut
        ActiveSheet.Paste
    End If
    If Range("A1").Value <> "Work_Day" Then Columns("A:A").Delete
    Application.DisplayAlerts = True
End Sub

