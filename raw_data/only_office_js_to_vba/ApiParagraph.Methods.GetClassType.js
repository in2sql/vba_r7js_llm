**Description / Описание**

English: This code retrieves the class type of a shape and inserts it into the document by adding a customized shape to the active worksheet in OnlyOffice.

Russian: Этот код получает тип класса формы и вставляет его в документ, добавляя настраиваемую форму на активный лист в OnlyOffice.

---

**OnlyOffice JavaScript Code / Код OnlyOffice JavaScript**

```javascript
// This example gets a class type and inserts it into the document.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified properties
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get the class type of the paragraph
var sClassType = oParagraph.GetClassType();

// Add text to the paragraph displaying the class type
oParagraph.AddText("Class Type = " + sClassType);
```

---

**Excel VBA Code / Код Excel VBA**

```vba
' This example gets a class type and inserts it into the document by adding a customized shape to the active worksheet.

Sub InsertClassType()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim sClassType As String
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet with specified properties
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph element
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Get the class type of the paragraph
    sClassType = oParagraph.GetClassType()
    
    ' Add text to the paragraph displaying the class type
    oParagraph.AddText "Class Type = " & sClassType
End Sub
```