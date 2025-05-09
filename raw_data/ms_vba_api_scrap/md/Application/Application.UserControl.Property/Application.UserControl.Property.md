# Application.UserControl property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
If Application.UserControl Then 
 MsgBox "This workbook was created by the user" 
Else 
 MsgBox "This workbook was created programmatically" 
End If
```

## Remarks
When the UserControl property is False for an object, that object is released when the last programmatic reference to the object is released. If this property is False, Microsoft Excel quits when the last object in the session is released.

## Example
No VBA example available.
