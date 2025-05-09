# PivotField PivotItems Method

## Business Description
Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the visible and hidden items (a PivotItems object) in the specified field. Read-only.

## Behavior
Returns an object that represents either a single PivotTable item (aPivotItemobject) or a collection of all the visible and hidden items (aPivotItemsobject) in the specified field. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtitem In pvtTable.PivotFields("product").PivotItemsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtitem.Name 
Next
```