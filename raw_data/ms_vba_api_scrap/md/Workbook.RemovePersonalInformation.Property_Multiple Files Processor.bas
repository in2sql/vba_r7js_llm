Attribute VB_Name = "JoKer"
Sub FilesProcessor()
    Dim xFd As FileDialog
    Dim xFdItem As Variant
    Dim xFileName As String
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set xFd = Application.FileDialog(msoFileDialogFolderPicker)
    If xFd.Show = -1 Then
        xFdItem = xFd.SelectedItems(1) & Application.PathSeparator
        xFileName = Dir(xFdItem & "*.xls*")
        Do While xFileName <> ""
        Set wb = Workbooks.Open(xFdItem & xFileName)
        DoWork wb
        wb.Close SaveChanges:=True
            xFileName = Dir
    Loop
    End If
    MsgBox "Done in shaa Allah"
    Application.ScreenUpdating = True
End Sub

Sub DoWork(wb As Workbook)
    With wb
' Do your work here
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
        ActiveSheet.ListObjects("Status").ShowHeaders = False
        Rows("1:1").Delete
    End If
    If Range("A1").Value <> "Work_Day" Then Columns("A:A").Delete
    Application.DisplayAlerts = True
' end of your code
End With
End Sub
