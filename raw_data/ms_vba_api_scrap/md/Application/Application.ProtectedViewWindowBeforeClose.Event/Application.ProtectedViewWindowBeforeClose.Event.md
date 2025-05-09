# Application.ProtectedViewWindowBeforeClose event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_ProtectedViewWindowBeforeClose(ByVal Pvw as ProtectedViewWindow, _ 
 Reason as XlProtectedViewCloseReason, Cancel as Boolean) 
 a = MsgBox("Do you really want to close the Protected View window?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```

## Parameters
- **Pvw**: Required
- **Reason**: Required
- **Cancel**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_ProtectedViewWindowBeforeClose(ByVal Pvw as ProtectedViewWindow, _ 
 Reason as XlProtectedViewCloseReason, Cancel as Boolean) 
 a = MsgBox("Do you really want to close the Protected View window?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```

