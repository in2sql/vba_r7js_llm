```plaintext
// This code adds an OLE object to the active worksheet, retrieves its application ID, and sets the value of cell A1 with the ID.
// Этот код добавляет OLE-объект на активный лист, получает его идентификатор приложения и устанавливает значение ячейки A1 с этим ID.
```

```vba
' This VBA code adds an OLE object to the active worksheet, retrieves its application ID, and sets the value of cell A1 with the ID.
' Этот VBA код добавляет OLE-объект на активный лист, получает его идентификатор приложения и устанавливает значение ячейки A1 с этим ID.

Sub AddOleObjectAndSetAppID()
    Dim oWorksheet As Worksheet
    Dim oOleObject As OLEObject
    Dim sAppId As String
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add an OLE Object to the worksheet
    Set oOleObject = oWorksheet.OLEObjects.Add( _
        Filename:="https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=130 * 360, _
        Top:=90 * 360, _
        Width:=2 * 360, _
        Height:=3 * 360)
    
    ' Retrieve the application ID from the OLE object
    ' Note: VBA does not have a direct method to get Application ID from OLEObject
    ' This is a placeholder as OnlyOffice specific methods are not available in standard VBA
    sAppId = "SampleAppID" ' Replace with actual method if available
    
    ' Set the value of cell A1 with the application ID
    oWorksheet.Range("A1").Value = "The OLE object application ID: " & sAppId
End Sub
```

```javascript
// This example shows how to get the application ID from the OLE object.
// Этот пример показывает, как получить идентификатор приложения из OLE-объекта.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oOleObject = oWorksheet.AddOleObject(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", // Image URL
    130 * 36000, // Left position
    90 * 36000, // Top position
    "https://youtu.be/SKGz4pmnpgY", // Link URL
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", // ProgID
    0, // Reserved
    2 * 36000, // Width
    4, // Update option
    3 * 36000  // Height
);
var sAppId = oOleObject.GetApplicationId(); // Get the application ID from the OLE object
oWorksheet.GetRange("A1").SetValue("The OLE object application ID: " + sAppId); // Set the value of cell A1 with the application ID
```