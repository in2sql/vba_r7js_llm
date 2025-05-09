# Workbook Deactivate Event

## Business Description
Occurs when the chart, worksheet, or workbook is deactivated.

## Behavior
Occurs when the chart, worksheet, or workbook is deactivated.

## Example Usage
```vba
Private Sub Workbook_Deactivate() 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```