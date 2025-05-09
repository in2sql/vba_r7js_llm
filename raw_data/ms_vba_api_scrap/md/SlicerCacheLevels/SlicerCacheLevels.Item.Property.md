# SlicerCacheLevels Item Property

## Business Description
Returns the specified SlicerCacheLevel object from the collection, or if no level is specified, returns the first SlicerCacheLevel object in the collection.

## Behavior
Returns the specifiedSlicerCacheLevelobject from the collection, or if no level is specified, returns the firstSlicerCacheLevelobject in the collection.

## Example Usage
```vba
ActiveWorkbook.SlicerCaches("Slicer_Country"). _ 
 SlicerCacheLevels("[Customer].[Customer Geography].[Country]")
```