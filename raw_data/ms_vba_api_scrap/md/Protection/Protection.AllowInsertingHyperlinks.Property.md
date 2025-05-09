# Protection AllowInsertingHyperlinks Property

## Business Description
Returns True if the insertion of hyperlinks is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the insertion of hyperlinks is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Unlock cell A1. 
 Range("A1").Locked = False 
 
 ' Allow hyperlinks to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingHyperlinks= False Then 
 ActiveSheet.Protect AllowInsertingHyperlinks:=True 
 End If 
 
 MsgBox "Hyperlinks can be inserted on this protected worksheet." 
 
End Sub
```