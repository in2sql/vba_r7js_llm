**Description / Описание:**

This code adds an OLE object to the active worksheet and sets its data to a specified URL.
Этот код добавляет OLE-объект на активный лист и устанавливает его данные на указанный URL.

```vba
' VBA Code to add an OLE Object and set its data

Sub AddOleObjectAndSetData()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add an OLE Object to the worksheet
    ' Parameters: Filename, Link, DisplayAsIcon, Left, Top, Width, Height
    Dim oOleObject As OLEObject
    Set oOleObject = oWorksheet.OLEObjects.Add( _
        Filename:="https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=130 * 36000, _
        Top:=90 * 36000, _
        Width:=2 * 36000, _
        Height:=3 * 36000)
    
    ' Set data for the OLE Object
    ' Assuming SetData is equivalent to setting the object’s source or linked data
    oOleObject.Object.Source = "https://youtu.be/eJxpkjQG6Ew"
End Sub
```

```javascript
// JavaScript Code to add an OLE Object and set its data

// This example sets the data to the current OLE object.
var oWorksheet = Api.GetActiveSheet();

// Add an OLE Object to the worksheet with specified parameters
var oOleObject = oWorksheet.AddOleObject(
    "https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", // Image URL
    130 * 36000, // Left position
    90 * 36000,  // Top position
    "https://youtu.be/SKGz4pmnpgY", // Source URL
    "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", // ProgID or similar identifier
    0, // Some parameter (e.g., Link)
    2 * 36000, // Width
    4, // Some parameter (e.g., DisplayAsIcon)
    3 * 36000  // Height
);

// Set data for the OLE Object
oOleObject.SetData("https://youtu.be/eJxpkjQG6Ew");
```