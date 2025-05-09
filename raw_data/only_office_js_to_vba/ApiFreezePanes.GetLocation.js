### Description / Описание

**English:**  
This code freezes the first column in the active sheet and writes the address of the frozen pane's location into cell B1, with "Location:" in A1.

**Russian:**  
Этот код замораживает первый столбец на активном листе и записывает адрес местоположения замороженной области в ячейку B1, с текстом "Location:" в A1.

---

### VBA Code

```vba
' This macro freezes the first column and writes the freeze pane location to the sheet
Sub FreezeFirstColumnAndWriteLocation()
    Dim ws As Worksheet
    Dim freezePane As Range
    Dim paneAddress As String
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Freeze the first column
    ws.Activate
    ws.Range("B1").Select
    ActiveWindow.FreezePanes = True
    
    ' Get the freeze pane location
    Set freezePane = ActiveWindow.FreezePanes
    paneAddress = freezePane.Address
    
    ' Write "Location:" in A1
    ws.Range("A1").Value = "Location: "
    
    ' Write the address of the freeze pane in B1
    ws.Range("B1").Value = paneAddress
End Sub
```

---

### OnlyOffice JS Code

```javascript
// This example freezes the first column and pastes the frozen range address into the table.
Api.SetFreezePanesType('column'); // Freeze the first column
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFreezePanes = oWorksheet.GetFreezePanes(); // Get the freeze panes object
var oRange = oFreezePanes.GetLocation(); // Get the location of the freeze panes
oWorksheet.GetRange("A1").SetValue("Location: "); // Set "Location:" in cell A1
oWorksheet.GetRange("B1").SetValue(oRange.GetAddress()); // Set the address of the frozen range in cell B1
```