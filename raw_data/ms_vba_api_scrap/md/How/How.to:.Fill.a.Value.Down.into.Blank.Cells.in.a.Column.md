# How to: Fill a Value Down into Blank Cells in a Column

## Business Description
The following example looks at column A, and if there is a blank cell, sets the value of the blank cell to equal the value of the cell above it. This continues down the entire column, filling a value down into each blank cell.

## Behavior
The following example looks at column A, and if there is a blank cell, sets the value of the blank cell to equal the value of the cell above it. This continues down the entire column, filling a value down into each blank cell.

## Example Usage
```vba
Sub FillCellsFromAbove()
    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    On Error Resume Next
    ' Look in column A
    With Columns(1)
        ' For blank cells, set them to equal the cell above
        .SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        'Convert the formula to a value
        .Value = .Value
    End With
    Err.Clear
    Application.ScreenUpdating = True
End Sub
```