# Get class type and insert it into the table / Получение типа класса и вставка его в таблицу

```javascript
// This example gets a class type and inserts it into the table.
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oColor = Api.CreateColorFromRGB(255, 111, 61); // Create a color with specified RGB values
oWorksheet.GetRange("A2").SetValue("Text with color"); // Set value in cell A2
oWorksheet.GetRange("A2").SetFontColor(oColor); // Set font color of cell A2
var sColorClassType = oColor.GetClassType(); // Get class type of the color
oWorksheet.GetRange("A4").SetValue("Class type = " + sColorClassType); // Insert class type into cell A4
```

```vba
' This example gets a class type and inserts it into the table.
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet() ' Get the active sheet
Dim oColor As Object
Set oColor = Api.CreateColorFromRGB(255, 111, 61) ' Create a color with specified RGB values
oWorksheet.Range("A2").Value = "Text with color" ' Set value in cell A2
oWorksheet.Range("A2").Font.Color = oColor ' Set font color of cell A2
Dim sColorClassType As String
sColorClassType = oColor.GetClassType() ' Get class type of the color
oWorksheet.Range("A4").Value = "Class type = " & sColorClassType ' Insert class type into cell A4
```