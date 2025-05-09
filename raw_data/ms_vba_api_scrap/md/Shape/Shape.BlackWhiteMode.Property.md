# Shape BlackWhiteMode Property

## Business Description
Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write MsoBlackWhiteModehttp://msdn.microsoft.com/library/2b4d7e22-1277-9f5c-ba52-a37e113477c1(Office.15).aspx.

## Behavior
Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/writeMsoBlackWhiteMode.

## Example Usage
```vba
Sub UseBlackWhiteMode() 
 
    Dim wksOne As Worksheet 
    Set wksOne = Application.Worksheets(1) 
    wksOne.Shapes(1).BlackWhiteMode= msoBlackWhiteGrayOutline 
 
End Sub
```