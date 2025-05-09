**Description: This code creates a blip fill to apply to a shape using the selected image as the object background.  
Описание: Этот код создает заливку blip для применения к форме, используя выбранное изображение в качестве фона объекта.**

```vba
' VBA Code
Sub AddShapeWithBlipFill()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartDatabase, 60, 35, 200, 100) ' Adjust dimensions as needed
    
    ' Set the fill using a picture
    With shp.Fill
        .UserPicture "https://api.onlyoffice.com/content/img/docbuilder/examples/icon_DocumentEditors.png"
        .TextureTile = msoTrue
    End With
    
    ' Remove the stroke
    shp.Line.Visible = msoFalse
End Sub
```

```javascript
// JavaScript Code
// This example creates a blip fill to apply to the object using the selected image as the object background.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateBlipFill("https://api.onlyoffice.com/content/img/docbuilder/examples/icon_DocumentEditors.png", "tile");
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```