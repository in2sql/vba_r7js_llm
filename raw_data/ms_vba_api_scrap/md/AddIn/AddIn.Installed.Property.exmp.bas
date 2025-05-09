Set a = AddIns("Solver Add-In") 
If a.Installed = True Then 
 MsgBox "The Solver add-in is installed" 
Else 
 MsgBox "The Solver add-in is not installed" 
End If