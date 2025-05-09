# Application.WindowDeactivate event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window) 
 Wn.WindowState = xlMinimized 
End Sub
```

## Parameters
- **Wb**: Required
- **Wn**: Required

## Remarks
For information about how to use event procedures with the Application object, see Using events with the Application object.

## Example
```vba
Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window) 
 Wn.WindowState = xlMinimized 
End Sub
```

