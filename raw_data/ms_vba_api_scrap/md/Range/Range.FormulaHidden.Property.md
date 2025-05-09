# Range FormulaHidden Property

## Business Description
Returns or sets a Variant value that indicates if the formula will be hidden when the worksheet is protected.

## Behavior
Returns or sets aVariantvalue that indicates if the formula will be hidden when the worksheet is protected.

## Example Usage
```vba
Sub HideFormulas() 
 
 Worksheets("Sheet1").Range("A1:B1").FormulaHidden= True 
 
End Sub
```