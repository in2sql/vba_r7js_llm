# Set Bullet or Numbering to a Paragraph / Установка маркера или нумерации для абзаца

```vba
' VBA Code to set bullet or numbering to a paragraph
Sub SetBulletToParagraph()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtBox As TextBox
    Dim para As ParagraphFormat
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStoredData, 120, 35, 200, 100)
    
    ' Add a textbox to the shape
    Set txtBox = shp.TextFrame
    txtBox.Characters.Text = " This is an example of the bulleted paragraph."
    
    ' Set bullet for the first paragraph
    With txtBox.Characters.Paragraphs(1).ParagraphFormat
        .Bullet.Visible = msoTrue
        .Bullet.Character = 8226 ' Unicode for bullet
        .Bullet.Font.Color.RGB = RGB(255, 111, 61)
    End With
End Sub
```

```javascript
// JavaScript Code to set bullet or numbering to the paragraph
function setBulletToParagraph() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Create a solid fill with RGB color
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    
    // Create a stroke with no fill
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    
    // Add a shape to the worksheet
    var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
    
    // Get the content of the shape
    var oDocContent = oShape.GetContent();
    
    // Get the first paragraph element
    var oParagraph = oDocContent.GetElement(0);
    
    // Create a bullet with "-" symbol
    var oBullet = Api.CreateBullet("-");
    
    // Set the bullet to the paragraph
    oParagraph.SetBullet(oBullet);
    
    // Add text to the paragraph
    oParagraph.AddText(" This is an example of the bulleted paragraph.");
}
```