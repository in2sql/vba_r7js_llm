# Workbook PivotTableOpenConnection Event

## Business Description
Occurs after a PivotTable report opens the connection to its data source.

## Behavior
Occurs after a PivotTable report opens the connection to its data source.

## Example Usage
```vba
Private Sub ConnectionApp_PivotTableOpenConnection(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been opened." 
 
End Sub
```