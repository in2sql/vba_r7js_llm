# PivotField Orientation Property

## Business Description
Returns or sets a XlPivotFieldOrientation value that represents the location of the field in the specified PivotTable report.

## Behavior
Returns or sets aXlPivotFieldOrientationvalue that represents the location of the field in the specified PivotTable report.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Set pvtField = pvtTable.PivotFields("ORDER_DATE") 
Select Case pvtField.OrientationCase xlHidden 
 MsgBox "Hidden field" 
 Case xlRowField 
 MsgBox "Row field" 
 Case xlColumnField 
 MsgBox "Column field" 
 Case xlPageField 
 MsgBox "Page field" 
 Case xlDataField 
 MsgBox "Data field" 
End Select
```