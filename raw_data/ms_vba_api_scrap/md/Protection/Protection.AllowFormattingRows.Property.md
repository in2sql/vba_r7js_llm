# Protection AllowFormattingRows Property

## Business Description
Returns True if the formatting of rows is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the formatting of rows is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow rows to be formatted on a protected worksheet. 
 If ActiveSheet.Protection.AllowFormattingRows= False Then 
 ActiveSheet.Protect AllowFormattingRows:=True 
 End If 
 
 MsgBox "Rows can be formatted on this protected worksheet." 
 
End Sub
```