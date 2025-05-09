Private Sub App_ProtectedViewWindowOpen(ByVal Pvw As ProtectedViewWindow) 
 MsgBox "You are opening the following workbook in Protected View: " _ 
 & Pvw.Caption 
End Sub