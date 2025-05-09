# SlicerPivotTables AddPivotTable Method

## Business Description
Adds a reference to a PivotTable to the SlicerPivotTables collection.

## Behavior
Adds a reference to a PivotTable to theSlicerPivotTablescollection.

## Example Usage
```vba
Dim pvts As SlicerPivotTables 
Set pvts = ActiveWorkbook.SlicerCaches("Slicer_Customer").PivotTables 
pvts.AddPivotTable(ActiveSheet.PivotTables("PivotTable1"))
```