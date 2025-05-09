# Protection Object

## Business Description
Represents the various types of protection options available for a worksheet.

## Behavior
Represents the various types of protection options available for a worksheet.

## Example Usage
```vba
Sub SetProtection() 
 
 Range("A1").Formula = "1" 
 Range("B1").Formula = "3" 
 Range("C1").Formula = "4" 
 ActiveSheet.Protect 
 
 ' Check the protection setting of the worksheet and act accordingly. 
 If ActiveSheet.Protection.AllowInsertingColumns = False Then 
 ActiveSheet.Protect AllowInsertingColumns:=True 
 MsgBox "Insert a column between 1 and 3" 
 Else 
 MsgBox "Insert a column between 1 and 3" 
 End If 
 
End Sub
```