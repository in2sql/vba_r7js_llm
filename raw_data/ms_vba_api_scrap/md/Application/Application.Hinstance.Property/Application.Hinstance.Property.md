# Application.Hinstance property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckHinstance() 
 
 MsgBox Application.Hinstance 
 
End Sub
```

## Remarks
Important
This property returns a correct handle only in the 32-bit version of Excel. In Excel, the HinstancePtr property was introduced, which works correctly in both 32-bit and 64-bit versions of Excel.

## Example
```vba
Sub CheckHinstance() 
 
 MsgBox Application.Hinstance 
 
End Sub
```

