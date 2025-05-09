# Font Superscript Property

## Business Description
True if the font is formatted as superscript; False by default. Read/write Variant.

## Behavior
Trueif the font is formatted as superscript;Falseby default. Read/writeVariant.

## Example Usage
```vba
n = Worksheets("Sheet1").Range("A1").Characters.Count 
Worksheets("Sheet1").Range("A1") _ 
 .Characters(n, 1).Font.Superscript= True
```