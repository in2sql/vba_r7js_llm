# GroupShapes Object

## Business Description
Represents the individual shapes within a grouped shape.

## Behavior
Represents the individual shapes within a grouped shape.

## Example Usage
```vba
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 10, 10, 100, 100).Name = "shpOne" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 150, 10, 100, 100).Name = "shpTwo" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 300, 10, 100, 100).Name = "shpThree" 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured msoTextureBlueTissuePaper 
 .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble 
 End With 
End With
```