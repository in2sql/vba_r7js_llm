# CubeField LayoutForm Property

## Business Description
Returns or sets the way the specified PivotTable items appear—in table format or in outline format. Read/write XlLayoutFormType.

## Behavior
Returns or sets the way the specified PivotTable items appear—in table format or in outline format. Read/writeXlLayoutFormType.

## Example Usage
```vba
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutForm= xlOutline 
 .LayoutSubtotalLocation = xlTop 
End With
```