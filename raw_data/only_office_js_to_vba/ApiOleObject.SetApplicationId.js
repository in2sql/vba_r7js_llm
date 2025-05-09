# Description / Описание

This example sets the application ID to the current OLE object.  
Этот пример устанавливает идентификатор приложения для текущего OLE объекта.

```javascript
// This example sets the application ID to the current OLE object
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet from the API
var oOleObject = oWorksheet.AddOleObject(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", // URL of the OLE object image
    130 * 36000, // Left position
    90 * 36000, // Top position
    "https://youtu.be/SKGz4pmnpgY", // Link to content
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", // ProgID
    0, // Display as icon
    2 * 36000, // Width
    4, // Additional parameter
    3 * 36000 // Height
);
oOleObject.SetApplicationId("asc.{E5773A43-F9B3-4E81-81D9-CE0A132470E7}"); // Set the application ID
```

```vba
' This example sets the application ID to the current OLE object
Sub SetApplicationId()
    Dim oWorksheet As Worksheet
    Dim oOleObject As OLEObject
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add OLE object to the worksheet
    Set oOleObject = oWorksheet.OLEObjects.Add( _
        ClassType:="asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", _
        Filename:="https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=130 * 36000, _
        Top:=90 * 36000, _
        Width:=2 * 36000, _
        Height:=3 * 36000)
    
    ' Set the application ID
    oOleObject.Object.ApplicationId = "asc.{E5773A43-F9B3-4E81-81D9-CE0A132470E7}"
End Sub
```