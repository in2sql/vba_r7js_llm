# Shape GroupItems Property

## Business Description
Returns a GroupShapes object that represents the individual shapes in the specified group. Use the Item method of the GroupShapes object to return a single shape from the group. Applies to Shape objects that represent grouped shapes. Read-only.

## Behavior
Returns aGroupShapesobject that represents the individual shapes in the specified group. Use theItemmethod of theGroupShapesobject to return a single shape from the group. Applies toShapeobjects that represent grouped shapes. Read-only.

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