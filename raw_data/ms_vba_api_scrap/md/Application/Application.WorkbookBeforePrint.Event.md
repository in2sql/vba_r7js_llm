# Application.WorkbookBeforePrint event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookBeforePrint(ByVal Wb As Workbook, _ 
 Cancel As Boolean) 
 For Each wk in Wb.Worksheets 
 wk.Calculate 
 Next 
End Sub
```

## Parameters
- **Wb**: Required
- **Cancel**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_WorkbookBeforePrint(ByVal Wb As Workbook, _ 
 Cancel As Boolean) 
 For Each wk in Wb.Worksheets 
 wk.Calculate 
 Next 
End Sub
```

