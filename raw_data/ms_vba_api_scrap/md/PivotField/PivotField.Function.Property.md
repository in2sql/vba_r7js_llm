# PivotField Function Property

## Business Description
Returns or sets the function used to summarize the PivotTable field (data fields only). Read/write XlConsolidationFunction.

## Behavior
Returns or sets the function used to summarize the PivotTable field (data fields only). Read/writeXlConsolidationFunction.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("Sum of 1994").Function= xlSum
```