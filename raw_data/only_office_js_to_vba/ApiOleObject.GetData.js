### Description / Описание

This example shows how to get the string data from the OLE object.
Этот пример демонстрирует, как получить строковые данные из OLE объекта.

```vba
' VBA code to get string data from the OLE object
Sub GetOleObjectData()
    Dim oWorksheet As Worksheet
    Dim oOleObject As OLEObject
    Dim sData As String
    
    ' Get the active sheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add an OLE Object with specified parameters
    Set oOleObject = oWorksheet.OLEObjects.Add( _
        ClassType:="https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=130 * 36000, _
        Top:=90 * 36000, _
        Width:=4 * 36000, _
        Height:=3 * 36000)
    
    ' Get data from the OLE object
    sData = oOleObject.Object.GetData()
    
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "The OLE object data: " & sData
End Sub
```

```javascript
// JavaScript code to get string data from the OLE object
function GetOleObjectData() {
    // Get the active sheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Add an OLE Object with specified parameters
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
    
    // Get data from the OLE object
    var sData = oOleObject.GetData();
    
    // Set value in cell A1
    oWorksheet.GetRange("A1").SetValue("The OLE object data: " + sData);
}
```