# Application.RecordRelative property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Worksheets("Sheet1").Activate 
If Application.RecordRelative = False Then 
 MsgBox ActiveCell.Address(ReferenceStyle:=xlA1) 
Else 
 MsgBox ActiveCell.Address(ReferenceStyle:=xlR1C1) 
End If
```

## Example
No VBA example available.
