Set wb = Workbooks("myaddin.xla")

' ===== Next Example =====

Set wb = Workbooks(AddIns("analysis toolpak").Name)

' ===== Next Example =====

On Error Resume Next ' turn off error checking 
Set wbMyAddin = Workbooks(AddIns("My Addin").Name) 
lastError = Err 
On Error Goto 0 ' restore error checking 
If lastError <> 0 Then 
 ' the add-in workbook isn't currently open. Manually open it. 
 Set wbMyAddin = Workbooks.Open(AddIns("My Addin").FullName) 
End If