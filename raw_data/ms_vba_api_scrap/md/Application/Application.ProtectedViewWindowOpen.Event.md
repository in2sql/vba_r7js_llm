# Application.ProtectedViewWindowOpen event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_ProtectedViewWindowOpen(ByVal Pvw As ProtectedViewWindow) 
 MsgBox "You are opening the following workbook in Protected View: " _ 
 & Pvw.Caption 
End Sub
```

## Parameters
- **Pvw**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_ProtectedViewWindowOpen(ByVal Pvw As ProtectedViewWindow) 
 MsgBox "You are opening the following workbook in Protected View: " _ 
 & Pvw.Caption 
End Sub
```

