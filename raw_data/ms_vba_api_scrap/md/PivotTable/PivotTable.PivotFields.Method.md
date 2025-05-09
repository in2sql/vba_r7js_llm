# PivotTable PivotFields Method

## Business Description
Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of both the visible and hidden fields (a PivotFields object) in the PivotTable report. Read-only.

## Behavior
Returns an object that represents either a single PivotTable field (aPivotFieldobject) or a collection of both the visible and hidden fields (aPivotFieldsobject) in the PivotTable report. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.PivotFieldsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```