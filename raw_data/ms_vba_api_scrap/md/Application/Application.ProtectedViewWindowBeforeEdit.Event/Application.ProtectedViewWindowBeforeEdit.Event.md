# Application.ProtectedViewWindowBeforeEdit event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_ProtectedViewWindowBeforeEdit(ByVal Pvw As ProtectedViewWindow, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really " _ 
 & "want to edit the workbook?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```

## Parameters
- **Pvw**: Required
- **Cancel**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_ProtectedViewWindowBeforeEdit(ByVal Pvw As ProtectedViewWindow, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really " _ 
 & "want to edit the workbook?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```

