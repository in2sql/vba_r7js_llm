# FormatCondition Modify Method

## Business Description
Modifies an existing conditional format.

## Behavior
Modifies an existing conditional format.

## Example Usage
```vba
Worksheets(1).Range("e1:e10").FormatConditions(1) _ 
 .ModifyxlCellValue, xlLess, "=$a$1"
```