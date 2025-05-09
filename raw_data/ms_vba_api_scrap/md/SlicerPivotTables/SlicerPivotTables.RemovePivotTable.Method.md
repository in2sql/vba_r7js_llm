# SlicerPivotTables RemovePivotTable Method

## Business Description
Removes a reference to a PivotTable from the SlicerPivotTables collection.

## Behavior
Removes a reference to a PivotTable from theSlicerPivotTablescollection.

## Example Usage
```vba
Dim pvts As SlicerPivotTables 
Set pvts = ActiveWorkbook.SlicerCaches("Slicer_Customer").PivotTables 
pvts.RemovePivotTable("PivotTable1")
```