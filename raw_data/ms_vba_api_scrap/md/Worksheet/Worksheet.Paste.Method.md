# Worksheet Paste Method

## Business Description
Pastes the contents of the Clipboard onto the sheet.

## Behavior
Pastes the contents of the Clipboard onto the sheet.

## Example Usage
```vba
Worksheets("Sheet1").Range("C1:C5").Copy 
ActiveSheet.PasteDestination:=Worksheets("Sheet1").Range("D1:D5")
```