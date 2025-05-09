# Range NumberFormatLocal Property

## Business Description
Returns or sets a Variant value that represents the format code for the object as a string in the language of the user.

## Behavior
Returns or sets aVariantvalue that represents the format code for the object as a string in the language of the user.

## Example Usage
```vba
MsgBox "The number format for cell A1 is " & _ 
 Worksheets("Sheet1").Range("A1").NumberFormatLocal
```