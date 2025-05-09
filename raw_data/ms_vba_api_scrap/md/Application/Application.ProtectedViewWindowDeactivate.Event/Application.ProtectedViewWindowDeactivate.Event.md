# Application.ProtectedViewWindowDeactivate event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_ProtectedViewWindowDeactivate(ByVal Pvw As ProtectedViewWindow) 
 Pvw.WindowState = xlProtectedViewWindowMinimized 
End Sub
```

## Parameters
- **Pvw**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_ProtectedViewWindowDeactivate(ByVal Pvw As ProtectedViewWindow) 
 Pvw.WindowState = xlProtectedViewWindowMinimized 
End Sub
```

