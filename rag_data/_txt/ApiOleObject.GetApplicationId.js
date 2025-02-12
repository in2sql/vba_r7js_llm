**Description / Описание:**

This script demonstrates how to add an OLE object to the active worksheet and retrieve its application ID.  
Этот скрипт демонстрирует, как добавить OLE-объект в активный лист и получить его идентификатор приложения.

```javascript
// This example shows how to get the application ID from the OLE object.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
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
); // Add OLE object with specified parameters
var sAppId = oOleObject.GetApplicationId(); // Get the application ID of the OLE object
oWorksheet.GetRange("A1").SetValue("The OLE object application ID: " + sAppId); // Set the application ID in cell A1
```

```vba
' This VBA script adds an OLE object to the active worksheet and retrieves its application ID.

Sub AddOleObjectAndGetAppId()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add the OLE object with specified parameters
    Dim oOleObject As OLEObject
    Set oOleObject = oWorksheet.OLEObjects.Add( _
        ClassType:="https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=130 * 36000, _
        Top:=90 * 36000, _
        Width:=2 * 36000, _
        Height:=3 * 36000 _
    )
    
    ' Get the Application ID of the OLE object
    Dim sAppId As String
    sAppId = oOleObject.Object.ApplicationID
    
    ' Set the Application ID in cell A1
    oWorksheet.Range("A1").Value = "The OLE object application ID: " & sAppId
End Sub
```