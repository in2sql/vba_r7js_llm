Private Sub App_ProtectedViewWindowBeforeEdit(ByVal Pvw As ProtectedViewWindow, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really " _ 
 & "want to edit the workbook?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub