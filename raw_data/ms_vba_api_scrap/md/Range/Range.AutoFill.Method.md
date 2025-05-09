# Range AutoFill Method

## Business Description
Performs an autofill on the cells in the specified range.

## Behavior
Performs an autofill on the cells in the specified range.

## Example Usage
```vba
Set sourceRange = Worksheets("Sheet1").Range("A1:A2") 
Set fillRange = Worksheets("Sheet1").Range("A1:A20") 
sourceRange.AutoFillDestination:=fillRange
```