Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, _ 
 ByVal Sh As Object) 
 Sh.Move After:=Wb.Sheets(Wb.Sheets.Count) 
End Sub