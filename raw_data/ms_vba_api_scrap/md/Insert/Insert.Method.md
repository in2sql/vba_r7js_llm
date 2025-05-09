# Insert Method

## Business Description
Inserts a cell or a range of cells into the datasheet and shifts other cells away to make space.

## Behavior
Inserts a cell or a range of cells into the datasheet and shifts other cells away to make space.

## Example Usage
```vba
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:C5").InsertShift:=xlShiftDown
```