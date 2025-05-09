**Description:**
Creates a complex color scheme by selecting from available schemes and adds a curved up arrow shape to the active worksheet.
Создает сложную цветовую схему, выбирая из доступных схем, и добавляет фигуру изогнутой стрелки вверх на активный лист.

**Excel VBA Code:**
```vba
' VBA code to create a complex color scheme and add a curved up arrow shape to the active worksheet

Sub AddCurvedUpArrow()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create scheme color "dk1"
    Dim oSchemeColor As Object
    Set oSchemeColor = Api.CreateSchemeColor("dk1")
    
    ' Create solid fill with the scheme color
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(oSchemeColor)
    
    ' Create stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a curved up arrow shape
    oWorksheet.AddShape "curvedUpArrow", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000
End Sub
```

**OnlyOffice JavaScript Code:**
```javascript
// JavaScript code to create a complex color scheme and add a curved up arrow shape to the active worksheet

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create scheme color "dk1"
var oSchemeColor = Api.CreateSchemeColor("dk1");

// Create solid fill with the scheme color
var oFill = Api.CreateSolidFill(oSchemeColor);

// Create stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a curved up arrow shape
oWorksheet.AddShape("curvedUpArrow", 60 * 36000, 35 * 36000, oFill, oStroke, 0, 2 * 36000, 1, 3 * 36000);
```