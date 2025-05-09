# SlicerCacheLevels Object

## Business Description
Represents the collection of hierarchy levels for the OLAP data source that is filtered by a slicer.

## Behavior
Represents the collection of hierarchy levels for the OLAP data source that is filtered by a slicer.

## Example Usage
```vba
ActiveWorkbook.SlicerCaches("Slicer_Customer_Geography"). _ 
 SlicerCacheLevels("[Customer].[Customer Geography].[Country]")
```