# Window SplitColumn Property

## Business Description
Returns or sets the column number where the window is split into panes (the number of columns to the left of the split line). Read/write Long.

## Behavior
Returns or sets the column number where the window is split into panes (the number of columns to the left of the split line). Read/writeLong.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitColumn= 1.5
```