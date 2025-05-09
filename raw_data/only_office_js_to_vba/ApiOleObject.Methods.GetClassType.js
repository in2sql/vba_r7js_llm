# Code Description

**English**:  
This code retrieves the active worksheet, adds an OLE object with specified parameters, obtains the class type of the added OLE object, and sets the value of cell A1 to display the class type.

**Russian**:  
Этот код получает активный лист, добавляет OLE-объект с указанными параметрами, получает тип класса добавленного OLE-объекта и устанавливает значение ячейки A1 для отображения типа класса.

```javascript
// This example gets a class type and inserts it into the document.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oOleObject = oWorksheet.AddOleObject(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", // URL of the image
    130 * 36000, // Width
    90 * 36000, // Height
    "https://youtu.be/SKGz4pmnpgY", // Link
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", // ProgID
    0, // Placement
    2 * 36000, // Left position
    4, // Top position
    3 * 36000  // Z order
);
var sType = oOleObject.GetClassType(); // Get the class type of the OLE object
oWorksheet.GetRange("A1").SetValue("Class type: " + sType); // Set cell A1 value
```

```vba
' This example gets a class type and inserts it into the document.
Sub InsertOLEObject()
    Dim oWorksheet As Object
    Dim oOleObject As Object
    Dim sType As String
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Add an OLE object to the worksheet
    Set oOleObject = oWorksheet.AddOleObject( _
        "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _ ' URL of the image
        130 * 36000, _ ' Width
        90 * 36000, _ ' Height
        "https://youtu.be/SKGz4pmnpgY", _ ' Link
        "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", _ ' ProgID
        0, _ ' Placement
        2 * 36000, _ ' Left position
        4, _ ' Top position
        3 * 36000)   ' Z order
    
    ' Get the class type of the OLE object
    sType = oOleObject.GetClassType()
    
    ' Set cell A1 value
    oWorksheet.GetRange("A1").SetValue "Class type: " & sType
End Sub
```