# Hyperlinks Add Method

## Business Description
Adds a hyperlink to the specified range or shape.

## Behavior
Adds a hyperlink to the specified range or shape.

## Example Usage
```vba
With Worksheets(1) 
 .Hyperlinks.AddAnchor:=.Range("a5"), _ 
 Address:="http://example.microsoft.com", _ 
 ScreenTip:="Microsoft Web Site", _ 
 TextToDisplay:="Microsoft" 
End With
```