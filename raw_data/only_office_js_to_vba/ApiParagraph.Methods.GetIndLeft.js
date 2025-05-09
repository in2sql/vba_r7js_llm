**Description / Описание**
This script demonstrates how to create a shape in the active worksheet, add paragraphs with specific left indentation, and retrieve the indentation value.
Этот скрипт демонстрирует, как создать фигуру на активном листе, добавить абзацы с заданным отступом слева и получить значение отступа.

```javascript
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

// Add text to the paragraph
oParagraph.AddText("This is a paragraph with the indent of 2 inches set to it. ");

// Add more text to the paragraph
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");

// Set left indentation to 2880
oParagraph.SetIndLeft(2880);

// Retrieve the left indentation value
var nIndLeft = oParagraph.GetIndLeft();

// Create a new paragraph
oParagraph = Api.CreateParagraph();

// Add text displaying the left indentation
oParagraph.AddText("Left indent: " + nIndLeft);

// Push the new paragraph to the document content
oDocContent.Push(oParagraph);
```

```vba
' Получить активный лист
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet()

' Создать сплошную заливку с RGB цветом (255, 111, 61)
Dim oFill As Object
Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))

' Создать обводку без заливки
Dim oStroke As Object
Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())

' Добавить фигуру на лист с заданными параметрами
Dim oShape As Object
Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)

' Получить содержимое фигуры
Dim oDocContent As Object
Set oDocContent = oShape.GetContent()

' Получить первый абзац
Dim oParagraph As Object
Set oParagraph = oDocContent.GetElement(0)

' Добавить текст к абзацу
oParagraph.AddText "This is a paragraph with the indent of 2 inches set to it. "

' Добавить дополнительные строки к абзацу
oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "

' Установить левый отступ в 2880
oParagraph.SetIndLeft 2880

' Получить значение левого отступа
Dim nIndLeft As Long
nIndLeft = oParagraph.GetIndLeft()

' Создать новый абзац
Set oParagraph = Api.CreateParagraph()

' Добавить текст с отображением левого отступа
oParagraph.AddText "Left indent: " & nIndLeft

' Добавить новый абзац к содержимому документа
oDocContent.Push oParagraph
```