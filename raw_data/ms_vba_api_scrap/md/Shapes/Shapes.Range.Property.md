# Shapes Range Property

## Business Description
Returns a ShapeRange object that represents a subset of the shapes in a Shapes collection.

## Behavior
Returns aShapeRangeobject that represents a subset of the shapes in aShapescollection.

## Example Usage
```vba
Set myDocument = Worksheets(1) 
myDocument.Shapes.Range(Array(1, 3)) _ 
 .Fill.Patterned msoPatternHorizontalBrick
```