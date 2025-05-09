**Description (English):**  
This script adds a shape to the active worksheet, inserts paragraphs with specific text, sets the paragraph alignment to right, and applies a right indentation of 2 inches. It also retrieves the right indentation value and displays it in a new paragraph.

**Описание (Русский):**  
Этот скрипт добавляет фигуру на активный лист, вставляет абзацы с определённым текстом, устанавливает выравнивание абзаца по правому краю и применяет правый отступ в 2 дюйма. Также он получает значение правого отступа и отображает его в новом абзаце.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage",
    120 * 36000, // Left position
    70 * 36000,  // Top position
    oFill,
    oStroke,
    0,           // Rotation
    2 * 36000,   // Width
    0,           // Height
    3 * 36000    // Z-order
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the right offset of 2 inches set to it. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");

// Set paragraph alignment to right
oParagraph.SetJc("right");

// Set right indentation to 2880 (twips)
oParagraph.SetIndRight(2880);

// Retrieve the right indentation value
var nIndRight = oParagraph.GetIndRight();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the right indentation value
oParagraph.AddText("Right indent: " + nIndRight);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```

```vba
' Get the active worksheet
Dim oWorksheet As Object
Set oWorksheet = Application.ActiveSheet

' Create a solid fill with RGB color (255, 111, 61)
Dim oFill As Object
Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))

' Create a stroke with no fill
Dim oStroke As Object
Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())

' Add a shape to the worksheet with specified parameters
Dim oShape As Object
Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)

' Get the content of the shape
Dim oDocContent As Object
Set oDocContent = oShape.GetContent()

' Get the first paragraph element
Dim oParagraph As Object
Set oParagraph = oDocContent.GetElement(0)

' Add text to the paragraph
Call oParagraph.AddText("This is a paragraph with the right offset of 2 inches set to it. ")
Call oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ")

' Set paragraph alignment to right
Call oParagraph.SetJc("right")

' Set right indentation to 2880 (twips)
Call oParagraph.SetIndRight(2880)

' Retrieve the right indentation value
Dim nIndRight As Long
nIndRight = oParagraph.GetIndRight()

' Create a new paragraph
Set oParagraph = Api.CreateParagraph()

' Add text displaying the right indentation value
Call oParagraph.AddText("Right indent: " & nIndRight)

' Push the new paragraph to the document content
Call oDocContent.Push(oParagraph)
```