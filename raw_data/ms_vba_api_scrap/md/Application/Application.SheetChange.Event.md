# Application.SheetChange event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub Workbook_SheetChange(ByVal Sh As Object, _ 
 ByVal Source As Range) 
 ' runs when a sheet is changed 
End Sub
```

## Parameters
- **Sh**: Required
- **Target**: Required

## Remarks
This event doesn't occur on chart sheets.

## Example
```vba
Private Sub Workbook_SheetChange(ByVal Sh As Object, _ 
 ByVal Source As Range) 
 ' runs when a sheet is changed 
End Sub
```

