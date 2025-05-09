## Description / Описание

**English:** This code sets up headers and data in the active worksheet, creates a pivot table in a new worksheet, adds specified row and data fields, and changes the position of a pivot field after a delay.

**Russian:** Этот код устанавливает заголовки и данные на активном листе, создает сводную таблицу на новом листе, добавляет определенные поля строк и данных, а также изменяет позицию поля сводной таблицы после задержки.

### Excel VBA Code

```vba
' This VBA code replicates the OnlyOffice JS functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add pivot table to new worksheet
    Dim ptSheet As Worksheet
    Set ptSheet = Worksheets.Add
    Dim ptCache As PivotCache
    Set ptCache = ThisWorkbook.PivotCaches.Create(xlDatabase, dataRange)
    Dim pt As PivotTable
    Set pt = ptCache.CreatePivotTable(TableDestination:=ptSheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add row fields
    pt.PivotFields("Region").Orientation = xlRowField
    pt.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pt.AddDataField pt.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set a cell value
    ptSheet.Range("A12").Value = "The Style field position will change soon"
    
    ' Change position of 'Style' field after delay
    Application.OnTime Now + TimeValue("00:00:05"), "ChangePivotFieldPosition"
End Sub

Sub ChangePivotFieldPosition()
    Dim pt As PivotTable
    Dim pf As PivotField
    Set pt = Worksheets("PivotTable1").PivotTables(1)
    Set pf = pt.PivotFields("Style")
    pf.Position = 1
End Sub
```

### OnlyOffice JS Code

```javascript
// This JavaScript code replicates the VBA functionality using OnlyOffice API

var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table into new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Set a cell value
var pivotWorksheet = Api.GetActiveSheet();
pivotWorksheet.GetRange('A12').SetValue('The Style field position will change soon');

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Change the position after 5 seconds
setTimeout(function () {
    pivotField.SetPosition(1);
}, 5000);
```