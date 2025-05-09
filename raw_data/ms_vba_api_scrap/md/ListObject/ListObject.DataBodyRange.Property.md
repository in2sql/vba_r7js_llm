# ListObject DataBodyRange Property

## Business Description
Returns a Range object that represents the range of values, excluding the header row, in a table. Read-only.

## Behavior
Returns aRangeobject that represents the range of values, excluding the header row, in a  table. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveWorksheet.ListObjects.Item(1).DataBodyRange.Select
```