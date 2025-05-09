# Comments Object

## Business Description
A collection of cell comments.

## Behavior
A collection of cell comments.

## Example Usage
```vba
With Worksheets(1).Range("e5").AddComment 
 .Visible = False 
 .Text "reviewed on " & Date 
End With
```