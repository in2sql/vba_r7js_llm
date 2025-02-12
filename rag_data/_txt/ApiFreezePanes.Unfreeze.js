# Description / Описание
This code example freezes the first column in a worksheet, then unfreezes all panes, retrieves the location of the freeze panes, and sets values in cells A1 and B1 to display this location.
Этот пример кода замораживает первый столбец на листе, затем разблокирует все панели, получает расположение замороженных панелей и устанавливает значения в ячейках A1 и B1 для отображения этого расположения.

```vba
' VBA Code to freeze first column, unfreeze panes, and display location
Sub ManageFreezePanes()
    Dim ws As Worksheet
    Dim freezeRange As Range
    Dim location As String
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Freeze the first column
    ws.Columns("B").Select
    ActiveWindow.FreezePanes = True
    
    ' Unfreeze all panes
    ActiveWindow.FreezePanes = False
    
    ' Get the location of freeze panes (assuming top-left cell)
    location = ws.Range("A1").Address
    
    ' Set values in A1 and B1
    ws.Range("A1").Value = "Location:"
    ws.Range("B1").Value = location
End Sub
```

```javascript
// JavaScript Code to freeze first column, unfreeze panes, and display location using OnlyOffice API
function manageFreezePanes(Api) {
    // Set freeze panes type to 'column'
    Api.SetFreezePanesType('column');
    
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Get the current freeze panes object
    var oFreezePanes = oWorksheet.GetFreezePanes();
    
    // Unfreeze all panes
    oFreezePanes.Unfreeze();
    
    // Get the location of freeze panes
    var oRange = oFreezePanes.GetLocation();
    
    // Set value in cell A1
    oWorksheet.GetRange("A1").SetValue("Location: ");
    
    // Set value in cell B1 with the location
    oWorksheet.GetRange("B1").SetValue(oRange + ""); 
}
```