# FormatConditions Object

## Business Description
Represents the collection of conditional formats for a single range.

## Behavior
Represents the collection of conditional formats for a single range.

## Example Usage
```vba
With Worksheets(1).Range("e1:e10").FormatConditions _ 
 .Add(xlCellValue, xlGreater, "=$a$1") 
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