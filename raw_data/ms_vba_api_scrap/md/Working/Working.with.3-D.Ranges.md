# Working with 3-D Ranges

## Business Description
If you are working with the same range on more than one sheet, use the Array function to specify two or more sheets to select. The following example formats the border of a 3-D range of cells.

## Behavior
If you are working with the same range on more than one sheet, use theArrayfunction to specify two or more sheets to select. The following example formats the border of a 3-D range of cells.

## Example Usage
```vba
Sub FormatSheets() 
 Sheets(Array("Sheet2", "Sheet3", "Sheet5")).Select 
 Range("A1:H1").Select 
 Selection.Borders(xlBottom).LineStyle = xlDouble 
End Sub
```