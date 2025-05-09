Sub SeeClipboard() 
 
 ' Determine if Office Clipboard can be displayed. 
 If Application.DisplayClipboardWindow = True Then 
 MsgBox "Office Clipboard can be displayed." 
 Else 
 MsgBox "Office Clipboard cannot be displayed." 
 End If 
 
End Sub