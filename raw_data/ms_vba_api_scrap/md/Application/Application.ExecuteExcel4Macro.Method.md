# Application.ExecuteExcel4Macro method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Worksheets("Sheet1").Activate 
Range("C3").Select 
MsgBox ExecuteExcel4Macro("GET.CELL(42)")
```

## Parameters
- **String**: Required

## Return Value
Variant

## Remarks
The Microsoft Excel 4.0 macro isn't evaluated in the context of the current workbook or sheet. This means that any references should be external and should specify an explicit workbook name. For example, to run the Microsoft Excel 4.0 macro "My_Macro" in Book1 you must use "Book1!My_Macro()". If you don't specify the workbook name, this method fails.

## Example
No VBA example available.
