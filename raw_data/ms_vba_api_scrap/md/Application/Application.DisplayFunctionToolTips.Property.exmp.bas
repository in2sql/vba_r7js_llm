Sub CheckToolTip() 
 
 ' Notify the user of the ability to display function ToolTips. 
 If Application.DisplayFunctionToolTips = True Then 
 MsgBox "The ability to display function ToolTips is on." 
 Else 
 MsgBox "The ability to display function ToolTips is off." 
 End If 
 
End Sub