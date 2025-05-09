# Application.WorkbookBeforeSave event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _ 
 ByVal SaveAsUI As Boolean, Cancel as Boolean) 
 a = MsgBox("Do you really want to save the workbook?", vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```

## Parameters
- **Wb**: Required
- **SaveAsUI**: Required
- **Cancel**: Required

## Return Value
Nothing

## Example
```vba
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _ 
 ByVal SaveAsUI As Boolean, Cancel as Boolean) 
 a = MsgBox("Do you really want to save the workbook?", vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```

