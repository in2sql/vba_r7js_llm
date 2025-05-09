# Application.GetCustomListContents method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
listArray = Application.GetCustomListContents(1) 
For i = LBound(listArray, 1) To UBound(listArray, 1) 
 Worksheets("sheet1").Cells(i, 1).Value = listArray(i) 
Next i
```

## Parameters
- **ListNum**: Required

## Return Value
Variant

## Example
No VBA example available.
