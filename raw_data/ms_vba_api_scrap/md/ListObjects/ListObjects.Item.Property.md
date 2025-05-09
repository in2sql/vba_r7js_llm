# ListObjects Item Property

## Business Description
Returns a single object from a collection.

## Behavior
Returns a single object from a collection.

## Example Usage
```vba
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects.Item(1).Name
```