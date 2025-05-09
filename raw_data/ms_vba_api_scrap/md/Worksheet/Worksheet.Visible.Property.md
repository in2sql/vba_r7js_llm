# Worksheet Visible Property

## Business Description
Returns or sets an XlSheetVisibility value that determines whether the object is visible.

## Behavior
Returns or sets anXlSheetVisibilityvalue that determines whether the object is visible.

## Example Usage
```vba
Set newSheet = Worksheets.Add 
newSheet.Visible= xlVeryHidden 
newSheet.Range("A1:D4").Formula = "=RAND()"
```