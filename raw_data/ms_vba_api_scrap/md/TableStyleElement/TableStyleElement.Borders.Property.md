# TableStyleElement Borders Property

## Business Description
Returns a Borders collection that represents the borders of a table style element. Read-only.

## Behavior
Returns aBorderscollection that represents the borders of a table style element. Read-only.

## Example Usage
```vba
With ActiveWorkbook.TableStyles("Table Style 4").TableStyleElements( _ 
 xlWholeTable).Borders(xlEdgeTop) 
 .Color = 255 
 .TintAndShade = 0 
 .Weight = 2 
 .LineStyle = 1 
End With
```