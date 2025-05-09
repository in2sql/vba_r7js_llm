# Range WrapText Property

## Business Description
Returns or sets a Variant value that indicates if Microsoft Excel wraps the text in the object.

## Behavior
Returns or sets aVariantvalue that indicates if Microsoft Excel wraps the text in the object.

## Example Usage
```vba
Worksheets("Sheet1").Range("B2").Value = _ 
 "This text should wrap in a cell." 
Worksheets("Sheet1").Range("B2").WrapText= True
```