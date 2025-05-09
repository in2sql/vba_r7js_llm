# PivotItem SourceName Property

## Business Description
Returns a Variant value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.

## Behavior
Returns aVariantvalue that represents the specified object's name as it appears in the original source data for the specified PivotTable report.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveSheet.PivotTables(1).PivotSelect "1998", xlDataAndLabel 
MsgBox "The original item name is " & _ 
 ActiveCell.PivotItem.SourceName
```