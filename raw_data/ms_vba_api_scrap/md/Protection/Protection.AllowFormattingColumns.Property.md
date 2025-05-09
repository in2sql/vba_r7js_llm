# Protection AllowFormattingColumns Property

## Business Description
Returns True if the formatting of columns is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the formatting of columns is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow columns to be formatted on a protected worksheet. 
 If ActiveSheet.Protection.AllowFormattingColumns= False Then 
 ActiveSheet.Protect AllowFormattingColumns:=True 
 End If 
 
 MsgBox "Columns can be formatted on this protected worksheet." 
 
End Sub
```