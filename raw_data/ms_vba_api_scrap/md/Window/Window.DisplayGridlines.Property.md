# Window DisplayGridlines Property

## Business Description
True if gridlines are displayed. Read/write Boolean.

## Behavior
Trueif gridlines are displayed. Read/writeBoolean.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayGridlines= Not(ActiveWindow.DisplayGridlines)
```