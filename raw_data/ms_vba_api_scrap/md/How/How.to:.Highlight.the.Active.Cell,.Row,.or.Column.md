# How to: Highlight the Active Cell, Row, or Column

## Business Description
The following code examples show ways to highlight the active cell or the rows and columns that contain the active cell. These examples use the SelectionChange event of the Worksheet object.

## Behavior
The following code examples show ways to highlight the active cell or the rows and columns that contain the active cell. These examples use theSelectionChangeevent of theWorksheetobject.

## Example Usage
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    ' Highlight the active cell
    Target.Interior.ColorIndex = 8
    Application.ScreenUpdating = True
End Sub
```