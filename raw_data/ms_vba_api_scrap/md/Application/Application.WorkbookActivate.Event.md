# Application.WorkbookActivate event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookActivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```

## Parameters
- **Wb**: Required

## Return Value
Nothing

## Remarks
For information about how to use event procedures with the Application object, see Using events with the Application object.

## Example
```vba
Private Sub App_WorkbookActivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```

