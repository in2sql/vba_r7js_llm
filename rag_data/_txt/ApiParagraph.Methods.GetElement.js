**Description / Описание**

This script adds a flowchart shape to the active worksheet, sets its fill and stroke properties, and adds three text runs to its first paragraph. The third text run is set to bold.

Этот скрипт добавляет фигуру блок-схемы на активный лист, устанавливает свойства заливки и обводки, а также добавляет три текстовых элемента в первый абзац. Третий текстовый элемент устанавливается жирным.

```javascript
// English: Add a flowchart shape to the active worksheet and modify its text runs.
// Russian: Добавить фигуру блок-схемы на активный лист и изменить ее текстовые элементы.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
// Add a flowchart shape with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.RemoveAllElements(); // Remove all existing elements in the paragraph

// Create and add the first text run
var oRun = Api.CreateRun();
oRun.AddText("This is the text for the first text run. Do not forget a space at its end to separate from the second one. ");
oParagraph.AddElement(oRun);

// Create and add the second text run
oRun = Api.CreateRun();
oRun.AddText("This is the text for the second run. We will set it bold afterwards. It also needs space at its end. ");
oParagraph.AddElement(oRun);

// Create and add the third text run
oRun = Api.CreateRun();
oRun.AddText("This is the text for the third run. It ends the paragraph.");
oParagraph.AddElement(oRun);

// Set the third text run to bold
oRun = oParagraph.GetElement(2);
oRun.SetBold(true); 
```

```vba
' English: Add a flowchart shape to the active worksheet and modify its text runs.
' Russian: Добавить фигуру блок-схемы на активный лист и изменить ее текстовые элементы.

Sub AddFlowChartShape()
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet
    
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)) ' Create a solid fill color
    
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill()) ' Create a stroke with no fill
    
    ' Add a flowchart shape with specified dimensions and styles
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent() ' Get the content of the shape
    
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0) ' Get the first paragraph element
    oParagraph.RemoveAllElements ' Remove all existing elements in the paragraph
    
    ' Create and add the first text run
    Dim oRun As Object
    Set oRun = Api.CreateRun()
    Call oRun.AddText("This is the text for the first text run. Do not forget a space at its end to separate from the second one. ")
    Call oParagraph.AddElement(oRun)
    
    ' Create and add the second text run
    Set oRun = Api.CreateRun()
    Call oRun.AddText("This is the text for the second run. We will set it bold afterwards. It also needs space at its end. ")
    Call oParagraph.AddElement(oRun)
    
    ' Create and add the third text run
    Set oRun = Api.CreateRun()
    Call oRun.AddText("This is the text for the third run. It ends the paragraph.")
    Call oParagraph.AddElement(oRun)
    
    ' Set the third text run to bold
    Set oRun = oParagraph.GetElement(2)
    Call oRun.SetBold(True)
End Sub
```