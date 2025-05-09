# PivotTable RepeatItemsOnEachPrintedPage Property

## Business Description
True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page. The default value is True. Read/write Boolean.

## Behavior
Trueif row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed.Falseif labels are printed only on the first page. The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable4") _ 
 .RepeatItemsOnEachPrintedPage= True
```