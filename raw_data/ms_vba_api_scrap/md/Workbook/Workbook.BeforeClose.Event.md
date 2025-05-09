# Workbook BeforeClose Event

## Business Description
Occurs before the workbook closes. If the workbook has been changed, this event occurs before the user is asked to save changes.

## Behavior
Occurs before the workbook closes. If the workbook has been changed, this event occurs before the user is asked to save changes.

## Example Usage
```vba
Private Sub Workbook_BeforeClose(Cancel as Boolean) 
 If Me.Saved = False Then Me.Save 
End Sub
```