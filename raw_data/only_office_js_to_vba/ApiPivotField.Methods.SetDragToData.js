**Description / Описание**

This code sets up a worksheet with Regions, Styles, and Prices, populates data, creates a pivot table on a new worksheet, configures row and column fields, adds a data field, adjusts the pivot field properties, and writes some values related to dragging.

Код настраивает лист с Регионми, Стилями и Ценами, заполняет данные, создаёт сводную таблицу на новом листе, настраивает строковые и столбцовые поля, добавляет поле данных, изменяет свойства поля сводной таблицы и записывает некоторые значения, связанные с перетаскиванием.

```vba
' VBA Code

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField

    ' Get the active sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set header values
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"

    ' Populate data
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

    ' Define the data range
    Set dataRange = ws.Range("B1:D5")

    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Insert a new worksheet for Pivot Table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"

    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="PivotTable1")

    ' Add row field
    With pivotTable.PivotFields("Style")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add column field
    With pivotTable.PivotFields("Region")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

    ' Get the PivotField for 'Region'
    Set pivotField = pivotTable.PivotFields("Region")

    ' Disable dragging to data area (Remove from data fields if present)
    On Error Resume Next
    pivotTable.DataFields(pivotField.Name).Orientation = xlHidden
    On Error GoTo 0

    ' Write some values in the pivot worksheet
    pivotWs.Range("A13").Value = "Drag to data"
    pivotWs.Range("B13").Value = False
    pivotWs.Range("A14").Value = "Try drag Region to data!"
End Sub
```

```javascript
// OnlyOffice JavaScript Code

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate data
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new Pivot Table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields
pivotTable.AddFields({
    rows: ['Style'],
    columns: 'Region',
});

// Add data field
pivotTable.AddDataField('Price');

// Get the active sheet (Pivot Table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Disable dragging to data area
pivotField.SetDragToData(false);

// Write some values in the pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to data');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToData());
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to data!');
```