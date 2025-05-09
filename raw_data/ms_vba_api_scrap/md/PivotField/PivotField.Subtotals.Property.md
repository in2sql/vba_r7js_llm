# PivotField Subtotals Property

## Business Description
Returns or sets subtotals displayed with the specified field. Valid only for nondata fields. Read/write Variant.

## Behavior
Returns or sets subtotals displayed with the specified field. Valid only for nondata fields. Read/writeVariant.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveCell.PivotField.Subtotals(2)= True
```