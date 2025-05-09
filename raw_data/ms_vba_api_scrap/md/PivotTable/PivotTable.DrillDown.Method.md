# PivotTable DrillDown Method

## Business Description
Enables you to drill down into the data within an OLAP or PowerPivot based cube hierarchy.

## Behavior
Enables you to drill down into the data within an OLAP or PowerPivot based cube hierarchy.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1").DrillDown ActiveSheet.PivotTables( _
      "PivotTable1").PivotFields("[Customer].[Customer Geography].[Country]"). _
      PivotItems("[Customer].[Customer Geography].[Country].&[Australia]"), _
      ActiveSheet.PivotTables("PivotTable1").PivotRowAxis.PivotLines(1)
```