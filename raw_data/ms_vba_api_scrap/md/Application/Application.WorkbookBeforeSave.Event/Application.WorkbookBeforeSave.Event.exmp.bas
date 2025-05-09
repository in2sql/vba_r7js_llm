Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _ 
 ByVal SaveAsUI As Boolean, Cancel as Boolean) 
 a = MsgBox("Do you really want to save the workbook?", vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub