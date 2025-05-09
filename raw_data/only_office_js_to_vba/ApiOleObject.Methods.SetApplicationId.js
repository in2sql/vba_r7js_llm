### Description / Описание
**English:** This code sets the application ID for the current OLE object in an OnlyOffice worksheet.

**Russian:** Этот код устанавливает идентификатор приложения для текущего OLE-объекта в рабочем листе OnlyOffice.

```vba
' VBA Code to set the application ID for the current OLE object in OnlyOffice

Sub SetApplicationID()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Add an OLE object with specified parameters
    Dim oOleObject As Object
    Set oOleObject = oWorksheet.AddOleObject( _
        "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
        130 * 36000, _
        90 * 36000, _
        "https://youtu.be/SKGz4pmnpgY", _
        "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", _
        0, _
        2 * 36000, _
        4, _
        3 * 36000)
    
    ' Set the application ID for the OLE object
    oOleObject.SetApplicationId "asc.{E5773A43-F9B3-4E81-81D9-CE0A132470E7}"
End Sub
```

```javascript
// JavaScript Code to set the application ID for the current OLE object in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add an OLE object with specified parameters
var oOleObject = oWorksheet.AddOleObject(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png",
    130 * 36000,
    90 * 36000,
    "https://youtu.be/SKGz4pmnpgY",
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}",
    0,
    2 * 36000,
    4,
    3 * 36000
);

// Set the application ID for the OLE object
oOleObject.SetApplicationId("asc.{E5773A43-F9B3-4E81-81D9-CE0A132470E7}");
```