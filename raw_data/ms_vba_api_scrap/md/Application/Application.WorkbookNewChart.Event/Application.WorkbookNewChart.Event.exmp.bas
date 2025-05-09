Private Sub App_NewChart(ByVal Wb As Workbook, _ 
 ByVal Ch As Chart) 
 MsgBox ("A new chart was added to: " & Wb.Name & " of type: " & Ch.Type) 
End Sub