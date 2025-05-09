Sub UseCheckAbort() 
 
 Dim rngSubtotal As Variant 
 Set rngSubtotal = Application.Range("A10") 
 
 ' Stop recalculation except for designated cell. 
 Application.CheckAbort KeepAbort:=rngSubtotal 
 
End Sub