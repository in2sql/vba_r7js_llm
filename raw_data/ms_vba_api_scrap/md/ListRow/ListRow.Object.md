# ListRow Object

## Business Description
Represents a row in a table. The ListRow object is a member of the ListRows collection.

## Behavior
Represents a row in a table. TheListRowobject is a member of theListRowscollection.

## Example Usage
```vba
Dim wrksht As Worksheet 
Dim oListRow As ListRow 
 
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
Set oListRow = wrksht.ListObjects(1).ListRows.Add
```