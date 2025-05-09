# Application.HinstancePtr property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckHinstance() 
    MsgBox Application.HinstancePtr 
End Sub
```

## Remarks
This property returns a correct handle in both the 32-bit and 64-bit versions of Excel. It extends the functionality of the Hinstance property of the Application object, which only works correctly in the 32-bit version of Excel.

## Example
```vba
Sub CheckHinstance() 
    MsgBox Application.HinstancePtr 
End Sub
```

