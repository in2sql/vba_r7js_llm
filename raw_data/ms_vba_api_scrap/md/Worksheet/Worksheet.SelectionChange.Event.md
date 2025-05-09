# Worksheet SelectionChange Event

## Business Description
Occurs when the selection changes on a worksheet.

## Behavior
Occurs when the selection changes on a worksheet.

## Example Usage
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
 With ActiveWindow 
 .ScrollRow = Target.Row 
 .ScrollColumn = Target.Column 
 End With 
End Sub
```