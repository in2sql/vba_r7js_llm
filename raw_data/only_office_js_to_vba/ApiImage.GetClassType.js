### Description / Описание

**English:**  
This code retrieves the active worksheet, adds an image to it, obtains the image's class type, sets the widths of the first two columns, and inserts the class type value into cells A1 and B1.

**Russian:**  
Этот код получает активный лист, добавляет в него изображение, получает тип класса изображения, устанавливает ширину первых двух столбцов и вставляет значение типа класса в ячейки A1 и B1.

---

#### Excel VBA Code

```vba
' This VBA code retrieves the active worksheet, adds an image, gets its class type,
' sets column widths, and inserts the class type into cells A1 and B1.

Sub InsertImageAndClassType()
    Dim ws As Worksheet
    Dim img As Shape
    Dim classType As String
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Add an image to the worksheet
    Set img = ws.Shapes.AddPicture( _
        Filename:="https://api.onlyoffice.com/content/img/docbuilder/examples/coordinate_aspects.png", _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, _
        Left:=60 * 36000, _
        Top:=35 * 36000, _
        Width:=0, _
        Height:=2 * 36000)
    
    ' Get the class type of the image
    classType = TypeName(img)
    
    ' Set column widths
    ws.Columns(1).ColumnWidth = 15
    ws.Columns(2).ColumnWidth = 10
    
    ' Insert values into cells
    ws.Range("A1").Value = "Class Type = "
    ws.Range("B1").Value = classType
End Sub
```

#### OnlyOffice JavaScript Code

```javascript
// This JavaScript code retrieves the active worksheet, adds an image, gets its class type,
// sets column widths, and inserts the class type into cells A1 and B1.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add an image to the worksheet
var oImage = oWorksheet.AddImage("https://api.onlyoffice.com/content/img/docbuilder/examples/coordinate_aspects.png", 
                                 60 * 36000, 35 * 36000, 0, 2 * 36000, 2, 3 * 36000);

// Get the class type of the image
var sClassType = oImage.GetClassType();

// Set column widths
oWorksheet.SetColumnWidth(0, 15);
oWorksheet.SetColumnWidth(1, 10);

// Insert values into cells
oWorksheet.GetRange("A1").SetValue("Class Type = ");
oWorksheet.GetRange("B1").SetValue(sClassType);
```