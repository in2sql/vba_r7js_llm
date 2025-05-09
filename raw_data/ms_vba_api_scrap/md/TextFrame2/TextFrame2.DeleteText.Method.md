# TextFrame2 DeleteText Method

## Business Description
Deletes the text from a text frame and all the associated text properties.

## Behavior
Deletes the text from a text frame and all the associated text properties.

## Example Usage
```vba
With ActiveSheet.Shapes(1).TextFrame2 
 If .HasText Then 
 .DeleteText ()
```