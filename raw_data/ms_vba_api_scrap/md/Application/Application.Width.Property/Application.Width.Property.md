# Application.Width property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
With ActiveWindow 
 .WindowState = xlNormal 
 .Top = 1 
 .Left = 1 
 .Height = Application.UsableHeight 
 .Width = Application.UsableWidth 
End With
```

## Remarks
If the window is minimized, Width is read-only and returns the width of the window icon.

## Example
No VBA example available.
