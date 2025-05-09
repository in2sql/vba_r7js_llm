Sub SettingToolTip() 
 
 ' Notify the user of the ability to display the Insert Options button. 
 If Application.DisplayInsertOptions = True Then 
 MsgBox "The ability to display the Insert Options button is on." 
 Else 
 MsgBox "The ability to display the Insert Options button is off." 
 End If 
 
End Sub