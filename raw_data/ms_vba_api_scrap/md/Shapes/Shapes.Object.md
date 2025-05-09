# Shapes Object

## Business Description
A collection of all the Shape objects on the specified sheet.

## Behavior
A collection of all theShapeobjects on the specified sheet.

## Example Usage
```vba
Set myDocument = Worksheets(1) 
myDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _ 
 msoPatternHorizontalBrick
```