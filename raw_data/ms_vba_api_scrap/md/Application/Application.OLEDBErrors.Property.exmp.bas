Set objEr = Application.OLEDBErrors.Item(1) 
MsgBox "The following error occurred:" & _ 
 objEr.ErrorString & " : " & objEr.SqlState