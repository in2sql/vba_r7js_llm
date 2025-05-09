# Range Cells Property

## Business Description
Returns a Range object that represents the cells in the specified range.

## Behavior
Returns aRangeobject that represents the cells in the specified range.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range(Cells(1, 1),Cells(5, 3)).Font.Italic = True
```