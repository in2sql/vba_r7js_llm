# Worksheet PasteSpecial Method

## Business Description
Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.

## Behavior
Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.

## Example Usage
```vba
Worksheets("Sheet1").Range("D1").Select 
ActiveSheet.PasteSpecialformat:= _ 
 "Microsoft Word 8.0 Document Object"
```