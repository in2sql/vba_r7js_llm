# Font Subscript Property

## Business Description
True if the font is formatted as subscript. False by default. Read/write Variant.

## Behavior
Trueif the font is formatted as subscript.Falseby default. Read/writeVariant.

## Example Usage
```vba
Worksheets("Sheet1").Range("A1") _ 
 .Characters(2, 1).Font.Subscript= True
```