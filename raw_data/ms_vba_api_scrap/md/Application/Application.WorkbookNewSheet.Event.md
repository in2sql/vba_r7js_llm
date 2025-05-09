# Application.WorkbookNewSheet event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, _ 
 ByVal Sh As Object) 
 Sh.Move After:=Wb.Sheets(Wb.Sheets.Count) 
End Sub
```

## Parameters
- **Wb**: Required
- **Sh**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, _ 
 ByVal Sh As Object) 
 Sh.Move After:=Wb.Sheets(Wb.Sheets.Count) 
End Sub
```

