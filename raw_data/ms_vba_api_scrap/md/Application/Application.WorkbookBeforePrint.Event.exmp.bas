Private Sub App_WorkbookBeforePrint(ByVal Wb As Workbook, _ 
 Cancel As Boolean) 
 For Each wk in Wb.Worksheets 
 wk.Calculate 
 Next 
End Sub