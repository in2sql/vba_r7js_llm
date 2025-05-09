# PivotCaches Object

## Business Description
Represents the collection of memory caches from the PivotTable reports in a workbook.

## Behavior
Represents the collection of memory caches from the PivotTable reports in a workbook.

## Example Usage
```vba
For Each pc In ActiveWorkbook.PivotCaches 
 pc.RefreshOnFileOpen = True 
Next
```