# Window WindowState Property

## Business Description
Returns or sets the state of the window. Read/write XlWindowState.

## Behavior
Returns or sets the state of the window. Read/writeXlWindowState.

## Example Usage
```vba
With ActiveWindow 
 .WindowState= xlNormal 
 .Top = 1 
 .Left = 1 
 .Height = Application.UsableHeight 
 .Width = Application.UsableWidth 
End With
```