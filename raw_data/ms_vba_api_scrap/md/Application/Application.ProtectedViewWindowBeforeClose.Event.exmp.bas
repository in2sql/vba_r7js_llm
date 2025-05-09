Private Sub App_ProtectedViewWindowBeforeClose(ByVal Pvw as ProtectedViewWindow, _ 
 Reason as XlProtectedViewCloseReason, Cancel as Boolean) 
 a = MsgBox("Do you really want to close the Protected View window?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub