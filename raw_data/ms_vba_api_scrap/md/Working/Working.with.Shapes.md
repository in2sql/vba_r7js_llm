# Working with Shapes

## Business Description
Shapes, or drawing objects, are represented by three different objects: the Shapes collection, the ShapeRange collection, and the Shape object.

## Behavior
Shapes, or drawing objects, are represented by three different objects: theShapescollection, theShapeRangecollection, and theShapeobject. In general, you use theShapescollection to create shapes and to iterate through all the shapes on a given worksheet; you use theShapeobject to format or modify a single shape; and you use theShapeRangecollection to modify multiple shapes the same way you work with multiple shapes in the user interface.

## Example Usage
```vba
Set myDocument = Worksheets(1) 
Set myRange = myDocument.Shapes.Range(Array("Big Star", _ 
 "Little Star")) 
myRange.Fill.PresetGradient _ 
 msoGradientHorizontal, 1, msoGradientBrass
```