# Workbook SheetBeforeDoubleClick Event

## Business Description
Occurs when any worksheet is double-clicked, before the default double-click action.

## Behavior
Occurs when any worksheet is double-clicked, before the default double-click action.

## Example Usage
```vba
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, _ 
 ByVal Target As Range, ByVal Cancel As Boolean) 
 Cancel = True 
End Sub
```