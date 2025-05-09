Attribute VB_Name = "Filter2txt"
Sub Filter2txt()
Attribute Filter2txt.VB_ProcData.VB_Invoke_Func = " \n14"
' Apply this macro on a copy of your final (ready) file
' Put your file in a new folder, where new files is to be generated there
' Let the to-be-portioned sheet be the only sheet
    Dim WbName As String
    Dim Ws As Worksheet
    ActiveWorkbook.RemovePersonalInformation = False
    ActiveWorkbook.Save
    Application.ScreenUpdating = False
    WbName = Application.ActiveWorkbook.Name
    If ActiveWorkbook.Worksheets.Count > 1 Then
        If IsEmpty(ActiveSheet.UsedRange) = False And ActiveSheet.Range("A1") = "SourceID" And ActiveSheet.Range("O1") <> "" And ActiveSheet.Range("A1").ListObject Is Nothing Then GoTo 14
        Worksheets(1).Activate
        For Each Ws In ActiveWorkbook.Worksheets
            Application.DisplayAlerts = False
            If IsEmpty(Ws.UsedRange) = False And Ws.Range("A1") = "SourceID" And Ws.Range("O1") <> "" And Ws.Range("A1").ListObject Is Nothing Then GoTo 7
            Ws.Delete
7            Application.DisplayAlerts = True
        Next Ws
    End If
14    If ActiveWorkbook.Worksheets.Count > 1 Then
        For Each Ws In ActiveWorkbook.Worksheets
            Application.DisplayAlerts = False
            If Ws.Name <> ActiveWorkbook.ActiveSheet.Name Then Ws.Delete
            Application.DisplayAlerts = True
        Next Ws
    End If
    On Error GoTo 24
    ActiveSheet.ShowAllData
24    If Range("O1").Value = "ExtractionDate" Or Range("O1").Value = "Extraction Date" Then Columns("P:Z").Delete
    If Range("P1").Value = "ExtractionDate" Or Range("P1").Value = "Extraction Date" Then Columns("Q:Z").Delete
    Columns("B:B").Copy
    Range("U1").Select
    ActiveSheet.Paste
    ActiveSheet.Columns("U:U").RemoveDuplicates Columns:=1, Header:=xlNo
    Range("U1").Delete Shift:=xlUp
    If Range("O1").Value = "ExtractionDate" Or Range("O1").Value = "Extraction Date" Then Columns("O:O").Copy
    If Range("P1").Value = "ExtractionDate" Or Range("P1").Value = "Extraction Date" Then Columns("P:P").Copy
    Range("V1").Select
    ActiveSheet.Paste
    ActiveSheet.Columns("V:V").RemoveDuplicates Columns:=1, Header:=xlNo
    Range("V1").Delete Shift:=xlUp
    For i = 1 To WorksheetFunction.CountA(Columns("U:U"))
        For j = 1 To WorksheetFunction.CountA(Columns("V:V"))
            Worksheets(1).Activate
            If Range("O1").Value = "ExtractionDate" Or Range("O1").Value = "Extraction Date" Then
                ActiveSheet.Range("A1").AutoFilter Field:=15, Criteria1:=Format(Cells(j, 22).Value, "yyyy-mm-dd")
                ActiveSheet.Range("A1").AutoFilter Field:=2, Criteria1:=Cells(i, 21).Value
                Range("A1").Select
                ActiveCell.CurrentRegion.Copy
                Sheets.Add After:=ActiveSheet
                ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
                Columns("O:O").NumberFormat = "yyyy-mm-dd"
                If Range("A2").Value <> "" Then
                    On Error Resume Next
                    ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.Path & "\" & Cells(2, 2).Value & ", " & Format(Cells(2, 15).Value, "yyyy-mm-dd") & ".txt", FileFormat:=xlText, CreateBackup:=False
                End If
            ElseIf Range("P1").Value = "ExtractionDate" Or Range("P1").Value = "Extraction Date" Then
                ActiveSheet.Range("A1").AutoFilter Field:=16, Criteria1:=Format(Cells(j, 22).Value, "yyyy-mm-dd")
                ActiveSheet.Range("A1").AutoFilter Field:=2, Criteria1:=Cells(i, 21).Value
                Range("A1").Select
                ActiveCell.CurrentRegion.Copy
                Sheets.Add After:=ActiveSheet
                ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
                Columns("P:P").NumberFormat = "yyyy-mm-dd"
                If Range("A2").Value <> "" Then
                    On Error Resume Next
                    ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.Path & "\" & Cells(2, 2).Value & ", " & Format(Cells(2, 16).Value, "yyyy-mm-dd") & ".txt", FileFormat:=xlText, CreateBackup:=False
                End If
            End If
                Application.DisplayAlerts = False
                ActiveSheet.Delete
                Application.DisplayAlerts = True
        Next j
    Next i
    Worksheets(1).Activate
    Columns("U:V").Delete
    On Error GoTo 27
    ActiveSheet.ShowAllData
27    Range("A1").Select
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.Path & "\" & Left(WbName, Len(WbName) - 5) & "_KZ", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    MsgBox "You can find the files in: " & Application.ActiveWorkbook.Path
    Application.ScreenUpdating = True
End Sub
