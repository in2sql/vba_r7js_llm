# TextFrame Characters Method

## Business Description
Returns a Characters object that represents a range of characters within a shape's text frame. You can use the Characters object to add and format characters within the text frame.

## Behavior
Returns aCharactersobject that represents a range of characters within a shape's text frame. You can use theCharactersobject to add and format characters within the text frame.

## Example Usage
```vba
With ActiveSheet.Shapes(1).TextFrame 
 .Characters.Text = "abcdefg" 
 .Characters(3, 1).Font.Bold = True 
End With
```