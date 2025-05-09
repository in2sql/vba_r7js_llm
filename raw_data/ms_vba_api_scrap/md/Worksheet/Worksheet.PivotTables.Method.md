# Worksheet PivotTables Method

## Business Description
Returns an object that represents either a single PivotTable report (a PivotTable object) or a collection of all the PivotTable reports (a PivotTables object) on a worksheet. Read-only.

## Behavior
Returns an object that represents either a single PivotTable report (aPivotTableobject) or a collection of all the PivotTable reports (aPivotTablesobject) on a worksheet. Read-only.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1"). _ 
 PivotFields("Sum of 1994").Function = xlSum
```