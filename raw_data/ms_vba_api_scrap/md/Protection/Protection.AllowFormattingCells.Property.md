# Protection AllowFormattingCells Property

## Business Description
Returns True if the formatting of cells is allowed on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the formatting of cells is allowed on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow cells to be formatted on a protected worksheet. 
 If ActiveSheet.Protection.AllowFormattingCells= False Then 
 ActiveSheet.Protect AllowFormattingCells:=True 
 End If 
 
 MsgBox "Cells can be formatted on this protected worksheet." 
 
End Sub
```