**Description:**
This code creates a blip fill using a selected image as the background for a shape and adds the shape to the active worksheet with specified dimensions and properties.

**Описание:**
Этот код создает заливку с изображением, используя выбранное изображение в качестве фона для формы, и добавляет форму на активный лист с заданными размерами и свойствами.

```vba
' VBA Code to create a blip fill and add a shape to the active worksheet
Sub AddShapeWithImageFill()
    Dim oWorksheet As Worksheet
    Dim oFill As FillFormat
    Dim oStroke As LineFormat
    Dim imgURL As String
    Dim shapeName As String
    Dim leftPos As Single, topPos As Single
    Dim width As Single, height As Single
    Dim rotation As Single, zOrder As Long, flip As MsoFlipCmd, msoRotate As MsoDirection

    ' Set the active worksheet
    Set oWorksheet = ActiveSheet

    ' Image URL for the fill
    imgURL = "https://api.onlyoffice.com/content/img/docbuilder/examples/icon_DocumentEditors.png"
    
    ' Add a shape to the worksheet
    With oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 100, 100)
        ' Apply picture fill
        .Fill.UserPicture (imgURL)
        .Fill.Tile = msoTrue ' Tile the image
        ' Remove the line (stroke)
        .Line.Visible = msoFalse
        ' Set rotation, z-order, flip, etc. if needed
        .Rotation = 0
        .ZOrder msoBringToFront
    End With
End Sub
```

```javascript
// JavaScript Code to create a blip fill and add a shape to the active sheet using OnlyOffice API
function addShapeWithImageFill() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a blip fill with the specified image URL and tile pattern
    var oFill = Api.CreateBlipFill("https://api.onlyoffice.com/content/img/docbuilder/examples/icon_DocumentEditors.png", "tile");
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape with the specified parameters to the worksheet
    oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
}
```