Sub CheckHyperlinks() 
 
 ' Determine if automatic formatting is enabled and notify user. 
 If Application.AutoFormatAsYouTypeReplaceHyperlinks = True Then 
 MsgBox "Automatic formatting for typing in hyperlinks is enabled." 
 Else 
 MsgBox "Automatic formatting for typing in hyperlinks is not enabled." 
 End If 
 
End Sub