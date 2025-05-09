# FormatCondition Borders Property

## Business Description
Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

## Behavior
Returns aBorderscollection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).

## Example Usage
```vba
Sub SetRangeBorder() 
 
 With Worksheets("Sheet1").Range("B2").Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 3 
 End With 
 
End Sub
```