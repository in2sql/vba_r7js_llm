# FormatCondition Object

## Business Description
Represents a conditional format.

## Behavior
Represents a conditional format.

## Example Usage
```vba
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
 With .Font 
 .Bold = True 
 .ColorIndex = 3 
 End With 
End With
```