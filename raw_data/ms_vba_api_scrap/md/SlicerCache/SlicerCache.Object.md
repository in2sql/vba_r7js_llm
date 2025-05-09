# SlicerCache Object

## Business Description
Represents the current filter state for a slicer and information about which PivotCache or WorkbookConnection the slicer is connected to.

## Behavior
Represents the current filter state for a slicer and information about whichPivotCacheorWorkbookConnectionthe slicer is connected to.

## Example Usage
```vba
With ActiveWorkbook 
 .SlicerCaches.Add("AdventureWorks", _ 
 "[Customer].[Customer Geography]").Slicers.Add SlicerDestination:="Sheet2", _ 
 Level:="[Customer].[Customer Geography].[Country]", Caption:="Country" 
End With
```