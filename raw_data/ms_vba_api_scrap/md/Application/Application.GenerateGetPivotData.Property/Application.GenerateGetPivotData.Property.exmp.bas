Sub PivotTableInfo() 
 
 ' Determine the ability to get PivotTable report data and notify user. 
 If Application.GenerateGetPivotData = True Then 
 MsgBox "The ability to get PivotTable report data is enabled." 
 Else 
 Msgbox "The ability to get PivotTable report data is disabled." 
 End If 
 
End Sub