# Range FormatConditions Property

## Business Description
Returns a FormatConditions collection that represents all the conditional formats for the specified range. Read-only.

## Behavior
Returns aFormatConditionscollection that represents all the conditional formats for the specified range. Read-only.

## Example Usage
```vba
Worksheets(1).Range("e1:e10").FormatConditions(1) _ 
 .ModifyxlCellValue, xlLess, "=$a$1"
```