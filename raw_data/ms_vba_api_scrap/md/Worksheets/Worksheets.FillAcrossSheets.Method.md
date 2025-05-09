# Worksheets FillAcrossSheets Method

## Business Description
Copies a range to the same area on all other worksheets in a collection.

## Behavior
Copies a range to the same area on all other worksheets in a collection.

## Example Usage
```vba
x = Array("Sheet1", "Sheet5", "Sheet7") 
Sheets(x).FillAcrossSheets_ 
 Worksheets("Sheet1").Range("A1:C5")
```