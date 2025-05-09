# PivotField ChildItems Property

## Business Description
Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the items (a PivotItems object) that are group children in the specified field, or children of the specified item. Read-only.

## Behavior
Returns an object that represents either a single PivotTable item (aPivotItemobject) or a collection of all the items (aPivotItemsobject) that are group children in the specified field, or children of the specified item. Read-only.

## Example Usage
```vba
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In _ 
 pvtTable.PivotFields("product") 
 .PivotItems("vegetables").ChildItemsrw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```