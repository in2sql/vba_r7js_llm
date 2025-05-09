# Worksheet PivotTableUpdate Event

## Business Description
Occurs after a PivotTable report is updated on a worksheet.

## Behavior
Occurs after a PivotTable report is updated on a worksheet.

## Example Usage
```vba
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been updated." 
 
End Sub
```