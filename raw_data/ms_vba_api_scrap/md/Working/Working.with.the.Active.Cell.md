# Working with the Active Cell

## Business Description
The ActiveCell property returns a Range object that represents the cell that is active. You can apply any of the properties or methods of a Range object to the active cell, as in the following example.

## Behavior
TheActiveCellproperty returns aRangeobject that represents the cell that is active. You can apply any of the properties or methods of aRangeobject to the active cell, as in the following example.

## Example Usage
```vba
Sub SetValue() 
 Worksheets("Sheet1").Activate 
 ActiveCell.Value = 35 
End Sub
```