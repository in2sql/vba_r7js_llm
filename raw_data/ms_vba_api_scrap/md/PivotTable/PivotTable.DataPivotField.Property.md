# PivotTable DataPivotField Property

## Business Description
Returns a PivotField object that represents all the data fields in a PivotTable. Read-only.

## Behavior
Returns aPivotFieldobject that represents all the data fields in a PivotTable. Read-only.

## Example Usage
```vba
Sub UseDataPivotField() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Move second PivotItem to the first position in PivotTable. 
 pvtTable.DataPivotField.PivotItems(2).Position = 1 
 
End Sub
```