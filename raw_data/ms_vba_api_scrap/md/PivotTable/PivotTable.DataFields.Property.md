# PivotTable DataFields Property

## Business Description
Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as data fields. Read-only.

## Behavior
Returns an object that represents either a single PivotTable field (aPivotFieldobject) or a collection of all the fields (aPivotFieldsobject) that are currently shown as data fields. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.DataFieldsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```