# PivotTable RowFields Property

## Business Description
Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently showing as row fields. Read-only.

## Behavior
Returns an object that represents either a single field in a PivotTable report (aPivotFieldobject) or a collection of all the fields (aPivotFieldsobject) that are currently showing as row fields. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.RowFieldsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```