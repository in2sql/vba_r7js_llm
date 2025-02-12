### Description / Описание
**English:** This code inserts an image into the active worksheet, retrieves its class type, sets the width of the first two columns, and displays the class type in cell B1.

**Русский:** Этот код вставляет изображение в активный лист, получает его тип класса, устанавливает ширину первых двух столбцов и отображает тип класса в ячейке B1.

```vba
' VBA Code: Inserts an image, retrieves its class type, sets column widths, and displays the class type.

Sub InsertImageAndClassType()
    ' Insert image into the active worksheet
    Dim oImage As Shape
    Set oImage = ActiveSheet.Shapes.AddPicture( _
        "https://api.onlyoffice.com/content/img/docbuilder/examples/coordinate_aspects.png", _
        msoFalse, msoCTrue, 60, 35, 200, 200) ' Left, Top, Width, Height in points

    ' Get the class type of the inserted image
    Dim sClassType As String
    sClassType = TypeName(oImage)

    ' Set the width of the first two columns
    Columns(1).ColumnWidth = 15
    Columns(2).ColumnWidth = 10

    ' Set values in cells A1 and B1
    Range("A1").Value = "Class Type = "
    Range("B1").Value = sClassType
End Sub
```

```javascript
// JavaScript Code: Inserts an image, retrieves its class type, sets column widths, and displays the class type.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add an image to the worksheet with specified position and size
var oImage = oWorksheet.AddImage(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/coordinate_aspects.png", 
    60 * 36000, 
    35 * 36000, 
    0, 
    2 * 36000, 
    2, 
    3 * 36000
);

// Get the class type of the inserted image
var sClassType = oImage.GetClassType();

// Set the width of the first two columns
oWorksheet.SetColumnWidth(0, 15);
oWorksheet.SetColumnWidth(1, 10);

// Set values in cells A1 and B1
oWorksheet.GetRange("A1").SetValue("Class Type = ");
oWorksheet.GetRange("B1").SetValue(sClassType); 
```