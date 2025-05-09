# Range AddIndent Property

## Business Description
Returns or sets a Variant value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)

## Behavior
Returns or sets aVariantvalue that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)

## Example Usage
```vba
With Worksheets("Sheet1").Range("A1") 
 .HorizontalAlignment = xlHAlignDistributed 
 .AddIndent = True 
End With
```