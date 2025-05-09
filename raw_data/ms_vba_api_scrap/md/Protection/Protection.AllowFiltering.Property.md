# Protection AllowFiltering Property

## Business Description
Returns True if the user is allowed to make use of an AutoFilter that was created before the sheet was protected. Read-only Boolean.

## Behavior
ReturnsTrueif the user is allowed to make use of an AutoFilter that was created before the sheet was protected. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Unlock row 1. 
 Rows("1:1").Locked = False 
 
 ' Allow row 1 to be filtered on a protected worksheet. 
 If ActiveSheet.Protection.AllowFiltering= False Then 
 ActiveSheet.Protect AllowFiltering:=True 
 End If 
 
 MsgBox "Row 1 can be filtered on this protected worksheet." 
 
End Sub
```