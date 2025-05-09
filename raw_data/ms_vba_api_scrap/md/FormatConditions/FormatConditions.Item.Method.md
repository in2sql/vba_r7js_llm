# FormatConditions Item Method

## Business Description
Returns a single object from a collection.

## Behavior
Returns a single object from a collection.

## Example Usage
```vba
With Worksheets(1).Range("e1:e10").FormatConditions.Item(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
End With
```