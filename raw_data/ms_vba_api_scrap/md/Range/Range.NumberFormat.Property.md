# Range NumberFormat Property

## Business Description
Returns or sets a Variant value that represents the format code for the object.

## Behavior
Returns or sets aVariantvalue that represents the format code for the object.

## Example Usage
```vba
Worksheets("Sheet1").Range("A17").NumberFormat = "General" 
Worksheets("Sheet1").Rows(1).NumberFormat = "hh:mm:ss" 
Worksheets("Sheet1").Columns("C"). _ 
 NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
```