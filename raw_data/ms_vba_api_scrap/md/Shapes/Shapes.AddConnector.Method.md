# Shapes AddConnector Method

## Business Description
Creates a connector. Returns a Shape object that represents the new connector. When a connector is added, it's not connected to anything.

## Behavior
Creates a connector. Returns aShapeobject that represents the new connector. When a connector is added, it's not connected to anything. Use theBeginConnectandEndConnectmethods to attach the beginning and end of a connector to other shapes in the document.

## Example Usage
```vba
Sub AddCanvasConnector() 
 
    Dim wksNew As Worksheet 
    Dim shpCanvas As Shape 
 
    Set wksNew = Worksheets.Add 
 
    'Add drawing canvas to new worksheet 
    Set shpCanvas = wksNew.Shapes.AddCanvas( _ 
        Left:=150, Top:=150, Width:=200, Height:=300) 
 
    'Add connector to the drawing canvas 
    shpCanvas.CanvasItems.AddConnector_ 
        Type:=msoConnectorStraight, BeginX:=150, _ 
        BeginY:=150, EndX:=200, EndY:=200 
 
End Sub
```