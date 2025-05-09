# Protection AllowDeletingColumns Property

## Business Description
Returns True if the deletion of columns is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the deletion of columns is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 'Unlock column A. 
 Columns("A:A").Locked = False 
 
 ' Allow column A to be deleted on a protected worksheet. 
 If ActiveSheet.Protection.AllowDeletingColumns= False Then 
 ActiveSheet.Protect AllowDeletingColumns:=True 
 End If 
 
 MsgBox "Column A can be deleted on this protected worksheet." 
 
End Sub
```