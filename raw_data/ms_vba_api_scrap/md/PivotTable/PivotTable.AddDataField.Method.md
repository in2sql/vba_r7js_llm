# PivotTable AddDataField Method

## Business Description
Adds a data field to a PivotTable report. Returns a PivotField object that represents the new data field.

## Behavior
Adds a data field to a PivotTable report. Returns aPivotFieldobject that represents the new data field.

## Example Usage
```vba
Sub AddMoreFields() 
 
 With ActiveSheet.PivotTables("PivotTable1") 
 .AddDataFieldActiveSheet.PivotTables( _ 
 "PivotTable1").PivotFields("Score"), "Total Score" 
 End With 
 
End Sub
```