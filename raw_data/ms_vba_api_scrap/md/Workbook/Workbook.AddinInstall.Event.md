# Workbook AddinInstall Event

## Business Description
Occurs when the workbook is installed as an add-in

## Behavior
Occurs when the workbook is installed as an add-in

## Example Usage
```vba
Private Sub Workbook_AddinInstall() 
 With Application.Commandbars("Standard").Controls.Add 
 .Caption = "The AddIn's menu item" 
 .OnAction = "'ThisAddin.xls'!Amacro" 
 End With End Sub 
End Sub
```