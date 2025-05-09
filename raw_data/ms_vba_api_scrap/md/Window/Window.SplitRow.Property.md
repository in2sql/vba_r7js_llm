# Window SplitRow Property

## Business Description
Returns or sets the row number where the window is split into panes (the number of rows above the split). Read/write Long.

## Behavior
Returns or sets the row number where the window is split into panes (the number of rows above the split). Read/writeLong.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitRow= 10
```