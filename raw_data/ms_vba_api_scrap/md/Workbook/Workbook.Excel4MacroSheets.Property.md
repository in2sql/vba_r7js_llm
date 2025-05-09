# Workbook Excel4MacroSheets Property

## Business Description
Returns a Sheets collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.

## Behavior
Returns aSheetscollection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.

## Example Usage
```vba
MsgBox "There are " & ActiveWorkbook.Excel4MacroSheets.Count & _ 
 " Microsoft Excel 4.0 macro sheets in this workbook."
```