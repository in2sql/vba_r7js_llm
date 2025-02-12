## Description

This code demonstrates how to add an OLE object to a worksheet, retrieve its data, and set a value in a specific cell with the OLE object data.  
Этот код демонстрирует, как добавить OLE объект на лист, получить его данные и установить значение в определенную ячейку с данными OLE объекта.

```javascript
// JavaScript Code using OnlyOffice API

// This example shows how to get the string data from the OLE object.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oOleObject = oWorksheet.AddOleObject(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", // Image URL
    130 * 36000, // Left position
    90 * 36000,  // Top position
    "https://youtu.be/SKGz4pmnpgY", // Data source URL
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", // ProgID
    0, // Display mode
    2 * 36000, // Width
    4,        // Height
    3 * 36000  // Additional parameter
);
var sData = oOleObject.GetData(); // Get data from the OLE object
oWorksheet.GetRange("A1").SetValue("The OLE object data: " + sData); // Set value in cell A1
```

```vba
' VBA Code equivalent using OnlyOffice API

' This example shows how to get the string data from the OLE object.
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet

Dim oOleObject As Object
Set oOleObject = oWorksheet.AddOleObject( _
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _ ' Image URL
    130 * 36000, _ ' Left position
    90 * 36000, _  ' Top position
    "https://youtu.be/SKGz4pmnpgY", _ ' Data source URL
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", _ ' ProgID
    0, _ ' Display mode
    2 * 36000, _ ' Width
    4, _        ' Height
    3 * 36000 _  ' Additional parameter
)

Dim sData As String
sData = oOleObject.GetData() ' Get data from the OLE object

oWorksheet.GetRange("A1").SetValue "The OLE object data: " & sData ' Set value in cell A1
```