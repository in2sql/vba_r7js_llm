# Protection AllowDeletingRows Property

## Business Description
Returns True if the deletion of rows is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the deletion of rows is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 'Unlock row 1. 
 Rows("1:1").Locked = False 
 
 ' Allow row 1 to be deleted on a protected worksheet. 
 If ActiveSheet.Protection.AllowDeletingRows= False Then 
 ActiveSheet.Protect AllowDeletingRows:=True 
 End If 
 
 MsgBox "Row 1 can be deleted on this protected worksheet." 
 
End Sub
```