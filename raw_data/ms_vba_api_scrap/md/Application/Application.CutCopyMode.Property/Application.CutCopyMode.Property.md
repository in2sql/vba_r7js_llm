# Application.CutCopyMode property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Select Case Application.CutCopyMode 
 Case Is = False 
 MsgBox "Not in Cut or Copy mode" 
 Case Is = xlCopy 
 MsgBox "In Copy mode" 
 Case Is = xlCut 
 MsgBox "In Cut mode" 
End Select
```

## Remarks
This example uses a message box to display the status of Cut or Copy mode.

## Example
No VBA example available.
