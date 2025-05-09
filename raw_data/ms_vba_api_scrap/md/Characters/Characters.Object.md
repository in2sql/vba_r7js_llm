# Characters Object

## Business Description
Represents characters in an object that contains text.

## Behavior
Represents characters in an object that contains text.

## Example Usage
```vba
With Worksheets("Sheet1").Range("B1") 
 .Value = "New Title" 
 .Characters(5, 5).Font.Bold = True 
End With
```