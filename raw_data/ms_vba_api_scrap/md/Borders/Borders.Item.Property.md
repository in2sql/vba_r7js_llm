# Borders Item Property

## Business Description
Returns a Border object that represents one of the borders of either a range of cells or a style.

## Behavior
Returns aBorderobject that represents one of the borders of either a range of cells or a style.

## Example Usage
```vba
Worksheets("Sheet1").Range("a1:g1"). _ 
 Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
```