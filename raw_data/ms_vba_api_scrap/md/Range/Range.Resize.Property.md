# Range Resize Property

## Business Description
Resizes the specified range. Returns a Range object that represents the resized range.

## Behavior
Resizes the specified range. Returns aRangeobject that represents the resized range.

## Example Usage
```vba
Set tbl = ActiveCell.CurrentRegion 
tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, _ 
 tbl.Columns.Count).Select
```