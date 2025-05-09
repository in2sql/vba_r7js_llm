**Description / Описание:**

This code retrieves a class type and inserts it into the document.
Этот код получает тип класса и вставляет его в документ.

```javascript
// This example gets a class type and inserts it into the document.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with specified RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var sClassType = oDocContent.GetClassType(); // Get the class type
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.SetJc("left"); // Set paragraph alignment to left
oParagraph.AddText("Class Type = " + sClassType); // Add text to the paragraph
```

```vba
' This example gets a class type and inserts it into the document.
Sub InsertClassType()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim sClassType As String
    Dim oParagraph As Object
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with specified RGB color
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the class type
    sClassType = oDocContent.GetClassType()
    
    ' Get the first paragraph element
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Set paragraph alignment to left
    oParagraph.SetJc "left"
    
    ' Add text to the paragraph
    oParagraph.AddText "Class Type = " & sClassType
End Sub
```