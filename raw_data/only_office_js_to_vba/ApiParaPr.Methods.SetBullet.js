# Set the bullet or numbering to the current paragraph
# Установить маркер или нумерацию для текущего абзаца

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Create a bullet with "-"
var oBullet = Api.CreateBullet("-");

// Set the bullet for the paragraph
oParaPr.SetBullet(oBullet);

// Add text to the paragraph
oParagraph.AddText(" This is an example of the bulleted paragraph.");
```

```vba
' Get the active worksheet
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet()

' Create a solid fill with RGB color (255, 111, 61)
Dim oFill As Object
Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))

' Create a stroke with no fill
Dim oStroke As Object
Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())

' Add a shape to the worksheet with specified parameters
Dim oShape As Object
Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)

' Get the content of the shape
Dim oDocContent As Object
Set oDocContent = oShape.GetContent()

' Get the first paragraph element
Dim oParagraph As Object
Set oParagraph = oDocContent.GetElement(0)

' Get the paragraph properties
Dim oParaPr As Object
Set oParaPr = oParagraph.GetParaPr()

' Create a bullet with "-"
Dim oBullet As Object
Set oBullet = Api.CreateBullet("-")

' Set the bullet for the paragraph
oParaPr.SetBullet oBullet

' Add text to the paragraph
oParagraph.AddText " This is an example of the bulleted paragraph."
```