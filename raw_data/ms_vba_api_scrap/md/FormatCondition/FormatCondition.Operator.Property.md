# FormatCondition Operator Property

## Business Description
Returns a Long value that represents the operator for the conditional format.

## Behavior
Returns aLongvalue that represents the operator for the conditional format.

## Example Usage
```vba
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 If .Operator= xlLess And .Formula1 = "5" Then 
 .Modify xlCellValue, xlBetween, "5", "15" 
 End If 
End With
```