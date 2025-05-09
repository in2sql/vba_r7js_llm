# PivotField ParentItems Property

## Business Description
Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the items (a PivotItems object) that are group parents in the specified field. The specified field must be a group parent of another field.

## Behavior
Returns an object that represents either a single PivotTable item (aPivotItemobject) or a collection of all the items (aPivotItemsobject) that are group parents in the specified field. The specified field must be a group parent of another field. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In pvtTable.PivotFields("product").ParentItemsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```