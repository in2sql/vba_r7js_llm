Sub UseAutoRecover() 
 
 Application.AutoRecover.Time = 5 
 
 MsgBox "The time that will elapse between each automatic " & _ 
 "save has been set to " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub