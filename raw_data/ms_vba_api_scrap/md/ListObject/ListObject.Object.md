# ListObject Object

## Business Description
Represents a ListObject object in the ListObjects collection.

## Behavior
Represents aListObject Object (Excel)object in theListObjectscollection.

## Example Usage
```vba
Dim wrksht As Worksheet 
Dim oListCol As ListRow 
 
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
Set oListCol = wrksht.ListObjects(1).ListRows.Add
```