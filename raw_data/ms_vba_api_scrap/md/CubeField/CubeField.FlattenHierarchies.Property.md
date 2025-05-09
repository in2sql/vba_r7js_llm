# CubeField FlattenHierarchies Property

## Business Description
Returns or sets whether items from all levels of hierarchies in a named set cube field are displayed in the same field of a PivotTable report based on an OLAP cube. Read/write

## Behavior
Returns or sets whether items from all levels of hierarchies in a named set cube field are displayed in the same field of a PivotTable report based on an OLAP cube. Read/write

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1").CubeFields("[Summary P&L]"). _FlattenHierarchies= True
```