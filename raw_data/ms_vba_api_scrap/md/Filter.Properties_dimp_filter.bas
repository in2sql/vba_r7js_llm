Attribute VB_Name = "Module113"
Sub dimp_filter()

Dim msheet As Worksheet
Dim lastRow As Long
lastRow = Range("A" & Rows.Count).End(xlUp).Row


' filter to new titles

    Selection.AutoFilter
    ActiveSheet.Columns("S:S").AutoFilter Field:=19, Criteria1:="#N/A"

' paste to new sheet

    Range("A1:S1" & lastRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    
End Sub

