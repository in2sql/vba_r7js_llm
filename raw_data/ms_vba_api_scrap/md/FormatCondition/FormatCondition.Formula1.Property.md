# FormatCondition Formula1 Property

## Business Description
Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-only String.

## Behavior
Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-onlyString.

## Example Usage
```vba
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 If .Operator = xlLess And .Formula1= "5" Then 
 .Modify xlCellValue, xlLess, "10" 
 End If 
End With
```