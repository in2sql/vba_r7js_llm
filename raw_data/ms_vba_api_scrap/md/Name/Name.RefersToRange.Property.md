# Name RefersToRange Property

## Business Description
Returns the Range object referred to by a Name object. Read-only.

## Behavior
Returns theRangeobject referred to by aNameobject. Read-only.

## Example Usage
```vba
p = Sheets(ActiveSheet.Name).Names("Print_Area").RefersToRange.Value 
MsgBox "Print_Area: " & UBound(p, 1) & " rows, " & _ 
 UBound(p, 2) & " columns"
```