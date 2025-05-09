# PivotTable DrillUp Method

## Business Description
Enables you to drill up into the data within an OLAP or PowerPivot based cube hierarchy.

## Behavior
Enables you to drill up into the data within an OLAP or PowerPivot based cube hierarchy.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1").DrillUp ActiveSheet.PivotTables( _
      "PivotTable1").PivotFields("[Customer].[Customer Geography].[Postal Code]"). _
      PivotItems( _
      "[Customer].[Customer Geography].[Postal Code].&[2450]&[Coffs Harbour]"), _
      ActiveSheet.PivotTables("PivotTable1").PivotRowAxis.PivotLines(1)
```