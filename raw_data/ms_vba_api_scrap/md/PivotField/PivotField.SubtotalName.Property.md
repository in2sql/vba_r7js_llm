# PivotField SubtotalName Property

## Business Description
Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string "Subtotal". Read/write String.

## Behavior
Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string "Subtotal". Read/writeString.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable2") _ 
 .PivotFields("state").SubtotalName= "Regional Subtotal"
```