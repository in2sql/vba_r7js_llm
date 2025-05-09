# Application.IsSandboxed property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckIfSandboxed(wbk As Workbook) 
 MsgBox wbk.Application.IsSandboxed 
End Sub
```

## Return Value
Boolean

## Remarks
Use the IsSandboxed property to determine if a workbook is open in a Protected View window.

## Example
```vba
Sub CheckIfSandboxed(wbk As Workbook) 
 MsgBox wbk.Application.IsSandboxed 
End Sub
```

