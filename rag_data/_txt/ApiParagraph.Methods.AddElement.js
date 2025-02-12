## Description / Описание

This code adds a shape with a text run to the active worksheet.
Этот код добавляет форму с текстовым запуском на активный рабочий лист.

### OnlyOffice JavaScript Code

```javascript
// This example adds a shape with a text run to the active worksheet.
var oWorksheet = Api.GetActiveSheet();
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
var oDocContent = oShape.GetContent();
var oParagraph = oDocContent.GetElement(0);
var oRun = Api.CreateRun();
oRun.AddText("This is just a sample text run. Nothing special.");
oParagraph.AddElement(oRun); 
```

### Excel VBA Code

```vba
' This example adds a shape with a text run to the active worksheet.
Sub AddShapeWithText()
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0)
    
    Dim oRun As Object
    Set oRun = Api.CreateRun()
    
    oRun.AddText "This is just a sample text run. Nothing special."
    
    oParagraph.AddElement oRun
End Sub
```