# Workbook Excel4IntlMacroSheets Property

## Business Description
Returns a Sheets collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.

## Behavior
Returns aSheetscollection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.

## Example Usage
```vba
MsgBox "There are " & _ 
 ActiveWorkbook.Excel4IntlMacroSheets.Count & _ 
 " Microsoft Excel 4.0 international macro sheets" & _ 
 " in this workbook."
```