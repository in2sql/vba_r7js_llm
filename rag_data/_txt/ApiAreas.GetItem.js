**Description / Описание**

English: This code retrieves the active worksheet, selects a range of cells, sets their values, retrieves a specific area from the selected range, and pastes it into another location.

Russian: Этот код получает активный лист, выбирает диапазон ячеек, устанавливает их значения, извлекает определенную область из выбранного диапазона и вставляет ее в другое место.

```vba
' VBA code to perform the same actions as the OnlyOffice JS example

Sub GetAndPasteRange()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1:D1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1:D1")
    
    ' Set the value to "1"
    oRange.Value = "1"
    
    ' Select the range
    oRange.Select
    
    ' Get the Areas collection
    Dim oAreas As Areas
    Set oAreas = oRange.Areas
    
    ' Get the first item from areas
    Dim oItem As Range
    Set oItem = oAreas.Item(1)
    
    ' Get the range A5
    Set oRange = oWorksheet.Range("A5")
    
    ' Set the value
    oRange.Value = "The first item from the areas: "
    
    ' Autofit the row height
    oRange.EntireRow.AutoFit
    
    ' Paste the oItem into B5
    oWorksheet.Range("B5").Value = oItem.Value
End Sub
```

```javascript
// OnlyOffice JS code to perform the same actions as the VBA example

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1:D1
var oRange = oWorksheet.GetRange("B1:D1");

// Set the value to "1"
oRange.SetValue("1");

// Select the range
oRange.Select();

// Get the Areas collection
var oAreas = oRange.GetAreas();

// Get the first item from areas
var oItem = oAreas.GetItem(1);

// Get the range A5
oRange = oWorksheet.GetRange('A5');

// Set the value
oRange.SetValue("The first item from the areas: ");

// Autofit the row height
oRange.AutoFit(false, true);

// Paste the oItem into B5
oWorksheet.GetRange('B5').Paste(oItem);
```