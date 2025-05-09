# Range CreateNames Method

## Business Description
Creates names in the specified range, based on text labels in the sheet.

## Behavior
Creates names in the specified range, based on text labels in the sheet.

## Example Usage
```vba
Set rangeToName = Worksheets("Sheet1").Range("A1:B3") 
rangeToName.CreateNamesLeft:=True
```