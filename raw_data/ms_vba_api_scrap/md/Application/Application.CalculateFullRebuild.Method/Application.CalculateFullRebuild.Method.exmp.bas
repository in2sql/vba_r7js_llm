Sub UseCalculateFullRebuild() 
 
 If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 Application.CalculateFullRebuild 
 End If 
 
End Sub