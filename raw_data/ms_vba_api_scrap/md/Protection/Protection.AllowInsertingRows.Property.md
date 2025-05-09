# Protection AllowInsertingRows Property

## Business Description
Returns True if the insertion of rows is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the insertion of rows is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow rows to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingRows= False Then 
 ActiveSheet.Protect AllowInsertingRows:=True 
 End If 
 
 MsgBox "Rows can be inserted on this protected worksheet." 
 
End Sub
```