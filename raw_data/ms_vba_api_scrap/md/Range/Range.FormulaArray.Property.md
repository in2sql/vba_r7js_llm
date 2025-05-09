# Range FormulaArray Property

## Business Description
Returns or sets the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returns null. Read/write Variant.

## Behavior
Returns or sets the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returnsnull. Read/writeVariant.

## Example Usage
```vba
Worksheets("Sheet1").Range("A1:C5").FormulaArray= "=3"
```