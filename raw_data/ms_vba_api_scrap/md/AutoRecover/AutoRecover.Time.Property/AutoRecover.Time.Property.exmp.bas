Sub SetTimeValue() 
 
 Application.AutoRecover.Time = 5 
 MsgBox "The AutoRecover time interval is set at " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub