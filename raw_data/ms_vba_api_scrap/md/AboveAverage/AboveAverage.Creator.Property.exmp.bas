Sub FindCreator() 
 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.Creator = &h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
 
End Sub