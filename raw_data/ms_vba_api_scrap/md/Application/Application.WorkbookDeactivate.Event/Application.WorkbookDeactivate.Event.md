# Application.WorkbookDeactivate event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```

## Parameters
- **Wb**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```

