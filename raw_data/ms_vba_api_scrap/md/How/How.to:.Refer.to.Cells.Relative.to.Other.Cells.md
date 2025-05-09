# How to: Refer to Cells Relative to Other Cells

## Business Description
A common way to work with a cell relative to another cell is to use the Offset property.

## Behavior
A common way to work with a cell relative to another cell is to use theOffsetproperty. In the following example, the contents of the cell that is one row down and three columns over from the active cell on the active worksheet are formatted as double-underlined.

## Example Usage
```vba
Sub Underline() 
 ActiveCell.Offset(1, 3).Font.Underline = xlDouble 
End Sub
```