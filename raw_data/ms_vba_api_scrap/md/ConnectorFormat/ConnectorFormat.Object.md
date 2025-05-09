# ConnectorFormat Object

## Business Description
Contains properties and methods that apply to connectors.

## Behavior
Contains properties and methods that apply to connectors.

## Example Usage
```vba
Set mainshape = ActiveWindow.Selection.ShapeRange(1) 
With mainshape 
 bx = .Left + .Width + 50 
 by = .Top + .Height + 50 
End With 
With ActiveSheet 
 For j = 1 To mainshape.ConnectionSiteCount 
 With .Shapes.AddConnector(msoConnectorStraight, _ 
 bx, by, bx + 50, by + 50) 
 .ConnectorFormat.EndConnect mainshape, j 
 .ConnectorFormat.Type = msoConnectorElbow 
 .Line.ForeColor.RGB = RGB(255, 0, 0) 
 l = .Left 
 t = .Top 
 End With 
 With .Shapes.AddTextbox(msoTextOrientationHorizontal, _ 
 l, t, 36, 14) 
 .Fill.Visible = False 
 .Line.Visible = False 
 .TextFrame.Characters.Text = j 
 End With 
 Next j 
End With
```