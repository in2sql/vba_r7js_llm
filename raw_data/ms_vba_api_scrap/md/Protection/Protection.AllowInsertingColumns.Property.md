# Protection AllowInsertingColumns Property

## Business Description
Returns True if the insertion of columns is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the insertion of columns is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow columns to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingColumns= False Then 
 ActiveSheet.Protect AllowInsertingColumns:=True 
 End If 
 
 MsgBox "Columns can be inserted on this protected worksheet." 
 
End Sub
```