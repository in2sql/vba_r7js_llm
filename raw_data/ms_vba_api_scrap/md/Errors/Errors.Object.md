# Errors Object

## Business Description
Represents the various spreadsheet errors for a range.

## Behavior
Represents the various spreadsheet errors for a range.

## Example Usage
```vba
Sub ErrorValue() 
 
 ' Place a number written as text in cell A1. 
 Range("A1").Formula = "'1" 
 
 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then 
 MsgBox "Cell A1 has a number as text." 
 Else 
 MsgBox "Cell A1 is a number." 
 End If 
 
End Sub
```