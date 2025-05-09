# PivotTable VisibleFields Property

## Business Description
Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the visible fields (a PivotFields object). Visible fields are shown as row, column, page or data fields. Read-only.

## Behavior
Returns an object that represents either a single field in a PivotTable report (aPivotFieldobject) or a collection of all the visible fields (aPivotFieldsobject). Visible fields are shown as row, column, page or data fields. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.VisibleFieldsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```