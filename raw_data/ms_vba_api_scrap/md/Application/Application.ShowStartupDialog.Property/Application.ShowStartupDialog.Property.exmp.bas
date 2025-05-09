Sub CheckStartupDialog() 
 
 ' Determine if the New Workbook task pane is enabled. 
 If Application.ShowStartupDialog = False Then 
 MsgBox "ShowStartupDialog is set to False." 
 Else 
 MsgBox "ShowStartupDialog is set to True." 
 End If 
 
End Sub