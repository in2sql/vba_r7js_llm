# Range Font Property

## Business Description
Returns a Font object that represents the font of the specified object.

## Behavior
Returns aFontobject that represents the font of the specified object.

## Example Usage
```vba
Sub CheckFont() 
 
 Range("A1").Select 
 
 ' Determine if the font name for selected cell is Arial. 
 If Range("A1").Font.Name = "Arial" Then 
 MsgBox "The font name for this cell is 'Arial'" 
 Else 
 MsgBox "The font name for this cell is not 'Arial'" 
 End If 
 
End Sub
```