Sub StillCalculating() 
 
 If Application.CalculationState = xlDone Then 
 MsgBox "Done" 
 Else 
 MsgBox "Not Done" 
 End If 
 
End Sub