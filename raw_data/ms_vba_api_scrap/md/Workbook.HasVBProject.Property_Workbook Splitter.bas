Attribute VB_Name = "WbSplit"
Sub SplitWorkbook()
'Updateby20140612
Dim FileExtStr As String
Dim FileFormatNum As Long
Dim xWs As Worksheet
Dim xWb As Workbook
Dim FolderName As String
Application.ScreenUpdating = False
Set xWb = Application.ActiveWorkbook
DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
FolderName = xWb.Path & "\" & xWb.Name & " " & DateString
MkDir FolderName
For Each xWs In xWb.Worksheets
    xWs.Copy
    If Val(Application.Version) < 12 Then
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat
            Case 51:
                FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56:
                FileExtStr = ".xls": FileFormatNum = 56
            Case Else:
                FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
    End If
    xFile = FolderName & "\" & Application.ActiveWorkbook.Sheets(1).Name & FileExtStr
    Application.ActiveWorkbook.SaveAs xFile, FileFormat:=FileFormatNum
    Application.ActiveWorkbook.Close False
Next
MsgBox "You can find the files in " & FolderName
Application.ScreenUpdating = True
End Sub
