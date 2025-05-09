# PivotField BaseItem Property

## Business Description
Returns or sets the item in the base field for a custom calculation. Valid only for data fields. Read/write Variant.

## Behavior
Returns or sets the item in the base field for a custom calculation. Valid only for data fields. Read/writeVariant.

## Example Usage
```vba
With Worksheets("Sheet1").Range("A3").PivotField 
 .Calculation = xlDifferenceFrom 
 .BaseField = "ORDER_DATE" 
 .BaseItem= "5/16/89" 
End With
```