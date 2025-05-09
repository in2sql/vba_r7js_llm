# Range Characters Property

## Business Description
Returns a Characters object that represents a range of characters within the object text. You can use the Characters object to format characters within a text string.

## Behavior
Returns aCharactersobject that represents a range of characters within the object text. You can use theCharactersobject to format characters within a text string.

## Example Usage
```vba
With Worksheets("Sheet1").Range("A1") 
 .Value = "abcdefg" 
 .Characters(3, 1).Font.Bold = True 
End With
```