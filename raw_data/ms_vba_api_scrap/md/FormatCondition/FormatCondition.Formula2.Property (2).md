# FormatCondition Formula2 Property

## Business Description
Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween.

## Behavior
Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional formatOperatorproperty isxlBetweenorxlNotBetween. Can be a constant value, a string value, a cell reference, or a formula. Read-onlyString.

## Example Usage
```vba
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 If .Operator = xlBetween And _ 
 .Formula1 = "5" And _ 
 .Formula2= "10" Then 
 .Modify xlCellValue, xlLess, "10" 
 End If 
End With
```