# CubeField HierarchizeDistinct Property

## Business Description
Returns or sets whether to order and remove duplicates when displaying the specified named set in a PivotTable report based on an OLAP cube. Read/write

## Behavior
Returns or sets whether to order and remove duplicates  when displaying the specified named set in a PivotTable report based on an OLAP cube. Read/write

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1").CubeFields("[Summary P&L]"). _HierarchizeDistinct= True
```