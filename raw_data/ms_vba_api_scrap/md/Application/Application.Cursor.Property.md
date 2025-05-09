# Application.Cursor property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub ChangeCursor() 
 
 Application.Cursor = xlIBeam 
 For x = 1 To 1000 
 For y = 1 to 1000 
 Next y 
 Next x 
 Application.Cursor = xlDefault 
 
End Sub
```

## Remarks
XlMousePointer can be one of these constants:

## Example
```vba
Sub ChangeCursor() 
 
 Application.Cursor = xlIBeam 
 For x = 1 To 1000 
 For y = 1 to 1000 
 Next y 
 Next x 
 Application.Cursor = xlDefault 
 
End Sub
```

