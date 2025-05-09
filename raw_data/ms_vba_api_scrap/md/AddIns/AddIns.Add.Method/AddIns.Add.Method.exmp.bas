Sub UseAddIn() 
 
 Set myAddIn = AddIns.Add(Filename:="A:\MYADDIN.XLA", _ 
 CopyFile:=True) 
 MsgBox myAddIn.Title & " has been added to the list" 
 
End Sub