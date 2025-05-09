# ShapeRange Object

## Business Description
Represents a shape range, which is a set of shapes on a document.

## Behavior
Represents a shape range, which is a set of shapes on a document.

## Example Usage
```vba
Set myDocument = Worksheets(1) 
myDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _ 
 msoPatternHorizontalBrick
```