# Description
This example retrieves a class type and inserts it into the document.
Этот пример получает тип класса и вставляет его в документ.

```vba
' This example retrieves a class type and inserts it into the document

Sub InsertClassType()
    Dim oWorksheet As Object
    Dim oOleObject As Object
    Dim sType As String
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Add an OLE Object to the worksheet
    Set oOleObject = oWorksheet.AddOleObject("https://api.onlyoffice.com/content/img/docbuilder/examples/ole-object-image.png", _
                                           130 * 36000, 90 * 36000, _
                                           "https://youtu.be/SKGz4pmnpgY", _
                                           "asc.{38E022EA-AD92-45FC-B22B-49DF39746DB4}", _
                                           0, 2 * 36000, 4, 3 * 36000)
    
    ' Get the class type of the OLE Object
    sType = oOleObject.GetClassType()
    
    ' Set the value of cell A1 to display the class type
    oWorksheet.GetRange("A1").SetValue "Class type: " & sType
End Sub
```

```javascript
// This example retrieves a class type and inserts it into the document.

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
); // Add an OLE Object
var sType = oOleObject.GetClassType(); // Get the class type
oWorksheet.GetRange("A1").SetValue("Class type: " + sType); // Insert class type into cell A1
```