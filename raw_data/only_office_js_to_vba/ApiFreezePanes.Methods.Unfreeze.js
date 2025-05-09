**Description / Описание:**

English: This code freezes the first column, then unfreezes all panes in the worksheet and displays the freeze pane location.

Russian: Этот код замораживает первый столбец, затем отменяет заморозку всех панелей на листе и отображает расположение замороженной панели.

```javascript
// JavaScript OnlyOffice API code example

// Freeze the first column
Api.SetFreezePanesType('column');

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the freeze panes object
var oFreezePanes = oWorksheet.GetFreezePanes();

// Unfreeze all panes
oFreezePanes.Unfreeze();

// Get the location of the freeze panes
var oRange = oFreezePanes.GetLocation();

// Set the value in cell A1
oWorksheet.GetRange("A1").SetValue("Location: ");

// Set the value in cell B1 with the location
oWorksheet.GetRange("B1").SetValue(oRange + ""); 
```

```vba
' Excel VBA equivalent code

Sub ManageFreezePanes()
    ' Freeze the first column
    ActiveWindow.SplitColumn = 1
    ActiveWindow.FreezePanes = True
    
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Get the freeze panes object
    ' Note: VBA does not have a direct equivalent to GetFreezePanes()
    ' To unfreeze, we set FreezePanes to False
    ActiveWindow.FreezePanes = False
    
    ' Get the location of the freeze panes
    ' VBA does not provide a direct method to get freeze pane location after unfreezing
    ' Here, we'll assume the location was A1 before unfreezing
    Dim freezeLocation As String
    freezeLocation = "A1"
    
    ' Set the value in cell A1
    ws.Range("A1").Value = "Location: "
    
    ' Set the value in cell B1 with the location
    ws.Range("B1").Value = freezeLocation
End Sub
```