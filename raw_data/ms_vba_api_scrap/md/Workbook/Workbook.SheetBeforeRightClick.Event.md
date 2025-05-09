# Workbook SheetBeforeRightClick Event

## Business Description
Occurs when any worksheet is right-clicked, before the default right-click action.

## Behavior
Occurs when any worksheet is right-clicked, before the default right-click action.

## Example Usage
```vba
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, _ 
 ByVal Target As Range, ByVal Cancel As Boolean) 
 Cancel = True 
End Sub
```