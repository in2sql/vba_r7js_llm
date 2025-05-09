Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " & Wb.Name & " workbook was successfully saved.") 
End If 
End Sub