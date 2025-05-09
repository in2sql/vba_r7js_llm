# Application.StandardFont property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
If Application.OperatingSystem Like "*Macintosh*" Then 
 Application.StandardFont = "Geneva" 
Else 
 Application.StandardFont = "Arial" 
End If
```

## Remarks
If you change the standard font by using this property, the change doesn't take effect until you restart Microsoft Excel.

## Example
No VBA example available.
