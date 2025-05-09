# Application.WorkbookBeforeClose event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookBeforeClose(ByVal Wb as Workbook, _ 
 Cancel as Boolean) 
 a = MsgBox("Do you really want to close the workbook?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```

## Parameters
- **Wb**: Required
- **Cancel**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_WorkbookBeforeClose(ByVal Wb as Workbook, _ 
 Cancel as Boolean) 
 a = MsgBox("Do you really want to close the workbook?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```

