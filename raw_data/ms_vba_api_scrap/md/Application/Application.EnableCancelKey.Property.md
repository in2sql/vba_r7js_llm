# Application.EnableCancelKey property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
On Error GoTo handleCancel 
Application.EnableCancelKey = xlErrorHandler 
MsgBox "This may take a long time: press ESC to cancel" 
For x = 1 To 1000000 ' Do something 1,000,000 times (long!) 
 ' do something here 
Next x 
 
handleCancel: 
If Err = 18 Then 
 MsgBox "You cancelled" 
End If
```

## Remarks
XlEnableCancelKey can be one of these constants:

## Example
No VBA example available.
