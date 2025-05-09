**Description / Описание**

This script adds a shape to the active worksheet, modifies its content by adding and then deleting a paragraph, and updates cell A9 to indicate that the paragraph has been removed.

Этот скрипт добавляет фигуру на активный лист, изменяет её содержимое, добавляя затем удаляя абзац, и обновляет ячейку A9, указывая на то, что абзац был удален.

---

**VBA Code**

```vba
' VBA Code to add a shape, manipulate its content, and update a cell

Sub ManipulateShape()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Define the fill color using RGB
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a flowchart shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(Type:=msoShapeFlowchartOfflineStorage, _
                                 Left:=60, Top:=35, Width:=100, Height:=70)
    
    ' Apply the solid fill color to the shape
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With
    
    ' Remove any existing text from the shape
    shp.TextFrame.Characters.Text = ""
    
    ' Add a new paragraph with sample text to the shape
    shp.TextFrame.Characters.Text = "This is just a sample text."
    
    ' Delete the paragraph by clearing the text
    shp.TextFrame.Characters.Text = ""
    
    ' Update cell A9 with a message indicating the paragraph was removed
    ws.Range("A9").Value = "The paragraph from the shape content was removed."
End Sub
```

---

**OnlyOffice JS Code**

```javascript
// This script adds a shape to the active worksheet, modifies its content by adding and deleting a paragraph, and updates cell A9.

var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with line style 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flowchart shape to the worksheet with specified dimensions and properties
var oShape = oWorksheet.AddShape("flowChartOfflineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 100 * 36000, 0, 70 * 36000);

// Get the content of the shape's document
var oDocContent = oShape.GetContent();

// Remove all existing elements from the shape's content
oDocContent.RemoveAllElements();

// Create a new paragraph
var oParagraph = Api.CreateParagraph();

// Add sample text to the paragraph
oParagraph.AddText("This is just a sample text.");

// Add the paragraph to the shape's document content
oDocContent.Push(oParagraph);

// Delete the paragraph from the shape's content
oParagraph.Delete();

// Update cell A9 with a message indicating the paragraph was removed
oWorksheet.GetRange("A9").SetValue("The paragraph from the shape content was removed.");
```