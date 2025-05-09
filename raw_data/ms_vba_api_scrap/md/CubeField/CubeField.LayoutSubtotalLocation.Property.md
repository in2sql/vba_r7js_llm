# CubeField LayoutSubtotalLocation Property

## Business Description
Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field. Read/write XlSubtototalLocationType.

## Behavior
Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field.  Read/writeXlSubtototalLocationType.

## Example Usage
```vba
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutForm = xlOutline 
 .LayoutSubtotalLocation= xlAtTop 
End With
```