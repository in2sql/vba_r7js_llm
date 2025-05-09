# Range NavigateArrow Method

## Business Description
Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a Range object that represents the new selection.

## Behavior
Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns aRangeobject that represents the new selection. This method causes an error if it's applied to a cell without visible tracer arrows.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A1").Select 
ActiveCell.NavigateArrowTrue, 1
```