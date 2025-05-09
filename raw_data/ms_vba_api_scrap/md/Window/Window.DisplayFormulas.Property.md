# Window DisplayFormulas Property

## Business Description
True if the window is displaying formulas; False if the window is displaying values. Read/write Boolean.

## Behavior
Trueif the window is displaying formulas;Falseif the window is displaying values. Read/writeBoolean.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayFormulas= True
```