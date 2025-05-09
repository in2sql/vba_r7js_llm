# Window DisplayHeadings Property

## Business Description
True if both row and column headings are displayed; False if no headings are displayed. Read/write Boolean.

## Behavior
Trueif both row and column headings are displayed;Falseif  no headings are displayed. Read/writeBoolean.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayHeadings= False
```