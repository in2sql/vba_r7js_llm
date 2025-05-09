# PivotField HiddenItems Property

## Business Description
Returns an object that represents either a single hidden PivotTable item (a PivotItem object) or a collection of all the hidden items (a PivotItems object) in the specified field. Read-only.

## Behavior
Returns an object that represents either a single hidden PivotTable item (aPivotItemobject) or a collection of all the hidden items (aPivotItemsobject) in the specified field. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In pvtTable.PivotFields("product").HiddenItemsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```