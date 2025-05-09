Sub CheckDisplayFeature() 
 
 ' Check if the options button can be displayed. 
 If Application.DisplayPasteOptions = True Then 
 MsgBox "The ability to display the Paste Options button is on." 
 Else 
 MsgBox "The ability to display the Paste Options button is off." 
 End If 
 
End Sub