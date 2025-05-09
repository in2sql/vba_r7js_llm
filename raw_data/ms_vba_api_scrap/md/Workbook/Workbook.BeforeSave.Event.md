# Workbook BeforeSave Event

## Business Description
Occurs before the workbook is saved.

## Behavior
Occurs before the workbook is saved.

## Example Usage
```vba
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, _ 
        Cancel as Boolean) 
    a = MsgBox("Do you really want to save the workbook?", vbYesNo) 
    If a = vbNo Then Cancel = True 
End Sub
```