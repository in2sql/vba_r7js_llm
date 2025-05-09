# PivotField AutoShow Method

## Business Description
Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report.

## Behavior
Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report.

## Example Usage
```vba
ActiveSheet.PivotTables("Pivot1").PivotFields("Company") _ 
 .AutoShowxlAutomatic, xlTop, 2, "Sum of Sales"
```