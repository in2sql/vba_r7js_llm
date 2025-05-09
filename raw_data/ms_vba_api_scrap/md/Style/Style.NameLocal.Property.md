# Style NameLocal Property

## Business Description
Returns or sets the name of the object, in the language of the user. Read-only String.

## Behavior
Returns or sets the name of the object, in the language of the user. Read-onlyString.

## Example Usage
```vba
With ActiveWorkbook.Styles(1) 
 MsgBox "The name of the style is " & .Name 
 MsgBox "The localized name of the style is " & .NameLocalEnd With
```