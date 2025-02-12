### Description / Описание

**English:**  
This script retrieves the active worksheet, creates a solid fill and stroke, adds a specific shape to the worksheet, obtains the content of the shape, retrieves the class type of the first element in the content, and inserts a text indicating the class type.

**Russian:**  
Этот скрипт получает активный лист, создает заливку и обводку, добавляет определенную фигуру на лист, извлекает содержимое фигуры, получает тип класса первого элемента в содержимом и вставляет текст с указанием типа класса.

---

### VBA Code / Код VBA

```vba
' Description: Retrieves the active worksheet, creates a solid fill and stroke,
' adds a specific shape, gets the content of the shape, retrieves the class type
' of the first element, and adds a text indicating the class type.

Sub AddShapeWithClassType()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Application.ActiveSheet

    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = oWorksheet.Shapes.Range().Fill
    oFill.ForeColor.RGB = RGB(255, 111, 61)
    oFill.Solid

    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = oWorksheet.Shapes.Range().Line
    oStroke.Visible = msoFalse

    ' Add a shape to the worksheet
    Dim oShape As Object
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 120, 70, 120, 70)
    oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)
    oShape.Line.Visible = msoFalse

    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.TextFrame

    ' Get the first paragraph
    Dim oParagraph As Object
    Set oParagraph = oDocContent.Paragraphs(1)

    ' Get the class type of the paragraph
    Dim sClassType As String
    sClassType = TypeName(oParagraph)

    ' Add text indicating the class type
    oParagraph.Text = "Class Type = " & sClassType
End Sub
```

---

### JavaScript Code / Код JavaScript

```javascript
// Description: Retrieves the active worksheet, creates a solid fill and stroke,
// adds a specific shape, gets the content of the shape, retrieves the class type
// of the first element, and adds a text indicating the class type.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the class type of the paragraph
var sClassType = oParagraph.GetClassType();

// Add text indicating the class type
oParagraph.AddText("Class Type = " + sClassType);
```