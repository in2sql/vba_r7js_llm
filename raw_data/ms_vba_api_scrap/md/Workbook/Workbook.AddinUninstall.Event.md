# Workbook AddinUninstall Event

## Business Description
Occurs when the workbook is uninstalled as an add-in.

## Behavior
Occurs when the workbook is uninstalled as an add-in.

## Example Usage
```vba
Private Sub Workbook_AddinUninstall() 
 Application.WindowState = xlMinimized 
End Sub
```