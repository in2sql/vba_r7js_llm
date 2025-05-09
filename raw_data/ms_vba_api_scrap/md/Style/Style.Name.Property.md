# Style Name Property

## Business Description
Returns a String value that represents the name of the object.

## Behavior
Returns  aStringvalue that represents the name of the object.

## Example Usage
```vba
With ActiveWorkbook.Styles(1) 
 MsgBox "The name of the style: " & .NameMsgBox "The localized name of the style: " & .NameLocal 
End With
```