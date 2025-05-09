Private Sub App_WorkbookBeforeClose(ByVal Wb as Workbook, _ 
 Cancel as Boolean) 
 a = MsgBox("Do you really want to close the workbook?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub