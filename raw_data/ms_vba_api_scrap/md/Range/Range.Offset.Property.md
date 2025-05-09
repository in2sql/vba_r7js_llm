# Range Offset Property

## Business Description
Returns a Range object that represents a range that's offset from the specified range.

## Behavior
Returns aRangeobject that represents a range that's offset from the specified range.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveCell.Offset(rowOffset:=3, columnOffset:=3).Activate
```