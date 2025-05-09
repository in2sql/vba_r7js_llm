# SlicerCaches Add2 Method

## Business Description
Adds a new SlicerCache object to the collection.

## Behavior
Adds a newSlicerCacheobject to the collection.

## Example Usage
```vba
ActiveWorkbook.SlicerCaches.Add(ActiveCell.PivotTable, _ 
 "[Customer].[Customer Geography]")
```