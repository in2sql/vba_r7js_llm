**English:**
The code retrieves the active sheet, creates a solid fill color, creates a stroke with no fill, adds a shape to the worksheet with specified dimensions and formatting, retrieves the parent sheet of the shape, gets the document content, accesses the first element, and adds text to the paragraph stating the parent sheet's name.

**Russian:**
Код получает активный лист, создает сплошную заливку цвета, создает обводку без заливки, добавляет форму на лист с указанными размерами и форматированием, получает родительский лист формы, получает содержимое документа, обращается к первому элементу и добавляет текст в абзац, указывающий имя родительского листа.

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Add a flow chart online storage shape to the worksheet
Dim oShape As Shape
Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 60, 35, 200, 150)

' Set the fill color of the shape
oShape.Fill.ForeColor.RGB = RGB(255, 111, 61)

' Set the line (stroke) properties of the shape
With oShape.Line
    .Weight = 0    ' Set line weight to 0
    .Visible = msoFalse    ' Make the line invisible (no fill)
End With

' Get the parent sheet of the shape
Dim oParentSheet As Worksheet
Set oParentSheet = oShape.Parent

' Add text to the shape's text frame
oShape.TextFrame.Characters.Text = "Parent sheet name is " & oParentSheet.Name
```

```javascript
// Get the active sheet
let oWorksheet = Api.GetActiveSheet();

// Create a solid fill color with RGB(255, 111, 61)
let oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with weight 0 and no fill
let oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a flow chart online storage shape to the worksheet with specified dimensions and formatting
let oShape = oWorksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the parent sheet of the shape
let oParentSheet = oShape.GetParentSheet();

// Get the document content of the shape
let oDocContent = oShape.GetDocContent();

// Access the first element (paragraph) in the document content
let oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph stating the parent sheet's name
oParagraph.AddText("Parent sheet name is " + oParentSheet.GetName());
```