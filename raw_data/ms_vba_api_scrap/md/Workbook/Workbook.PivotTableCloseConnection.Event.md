# Workbook PivotTableCloseConnection Event

## Business Description
Occurs after a PivotTable report closes the connection to its data source.

## Behavior
Occurs after a PivotTable report closes the connection to its data source.

## Example Usage
```vba
Private Sub ConnectionApp_PivotTableCloseConnection(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```