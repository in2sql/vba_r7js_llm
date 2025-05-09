# PivotField AutoSort Method

## Business Description
Establishes automatic field-sorting rules for PivotTable reports.

## Behavior
Establishes automatic field-sorting rules for PivotTable reports.

## Example Usage
```vba
ActiveSheet.PivotTables(1).PivotField("Company") _ 
 .AutoSortxlDescending, "Sum of Sales"
```