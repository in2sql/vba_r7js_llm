# Range SpecialCells Method

## Business Description
Returns a Range object that represents all the cells that match the specified type and value.

## Behavior
Returns aRangeobject that represents all the cells that match the specified type and value.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate
```