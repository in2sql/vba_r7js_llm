# Description
Sets the data to the current OLE object.
Устанавливает данные для текущего OLE-объекта.

```vba
' VBA code to add an OLE Object and set its data
Sub AddOleObject()
    Dim ws As Worksheet
    Dim oleObj As OLEObject

    ' Get the active sheet
    Set ws = ActiveSheet

    ' Add an OLE object with specified properties
    Set oleObj = ws.OLEObjects.Add(Filename:="https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
                                   Link:=False, _
                                   DisplayAsIcon:=False, _
                                   Left:=130, _       ' Position from the left
                                   Top:=90, _         ' Position from the top
                                   Width:=200, _      ' Width of the OLE object
                                   Height:=150)       ' Height of the OLE object

    ' Set data for the OLE object
    oleObj.Object.Document.URL = "https://youtu.be/eJxpkjQG6Ew"
End Sub
```

```javascript
// This example sets the data to the current OLE object.
var oWorksheet = Api.GetActiveSheet();
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
oOleObject.SetData("https://youtu.be/eJxpkjQG6Ew"); 
```