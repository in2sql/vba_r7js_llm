# Workbook SheetPivotTableUpdate Event

## Business Description
Occurs after the sheet of the PivotTable report has been updated.

## Behavior
Occurs after the sheet of the PivotTable report has been updated.

## Example Usage
```vba
Private Sub ConnectionApp_SheetPivotTableUpdate(ByVal shOne As Object, Target As PivotTable) 
 
 MsgBox "The SheetPivotTable connection has been updated." 
 
End Sub
```