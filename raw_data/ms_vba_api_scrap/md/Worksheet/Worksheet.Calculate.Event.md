# Worksheet Calculate Event

## Business Description
Occurs after the worksheet is recalculated, for the Worksheet object.

## Behavior
Occurs after the worksheet is recalculated, for theWorksheetobject.

## Example Usage
```vba
Private Sub Worksheet_Calculate() 
 Columns("A:F").AutoFit 
End Sub
```