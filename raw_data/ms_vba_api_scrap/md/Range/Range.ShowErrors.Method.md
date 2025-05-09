# Range ShowErrors Method

## Business Description
Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.

## Behavior
Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
If IsError(ActiveCell.Value) Then 
 ActiveCell.ShowErrorsEnd If
```