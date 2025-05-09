# Window Split Property

## Business Description
True if the window is split. Read/write Boolean.

## Behavior
Trueif the window is split. Read/writeBoolean.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
With ActiveWindow 
 .SplitColumn = 2 
 .SplitRow = 2 
End With
```