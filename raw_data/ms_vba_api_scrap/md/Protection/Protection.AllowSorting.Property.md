# Protection AllowSorting Property

## Business Description
Returns True if the sorting option is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the sorting option is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Unlock cells A1 through B5. 
 Range("A1:B5").Locked = False 
 
 ' Allow sorting to be performed on the protected worksheet. 
 If ActiveSheet.Protection.AllowSorting= False Then 
 ActiveSheet.Protect AllowSorting:=True 
 End If 
 
 MsgBox "For cells A1 through B5, sorting can be performed on the protected worksheet." 
 
End Sub
```