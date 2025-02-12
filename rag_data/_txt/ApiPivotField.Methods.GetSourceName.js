# Description / Описание

This script initializes a worksheet with data, creates a pivot table, modifies pivot fields, and sets values in specific cells.

Этот скрипт инициализирует лист данными, создает сводную таблицу, изменяет поля сводной таблицы и устанавливает значения в определенные ячейки.

## Excel VBA Equivalent Code / Эквивалентный код Excel VBA

```vba
Sub CreatePivot()
    ' Initialize the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

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

    ' Add a new worksheet for the pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"

    ' Create Pivot Cache
    Dim pCache As PivotCache
    Set pCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    ' Create Pivot Table
    Dim pTable As PivotTable
    Set pTable = pCache.CreatePivotTable(TableDestination:=pivotWs.Range("A1"), TableName:="PivotTable1")

    ' Add row fields
    With pTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField

        ' Set row layout to Tabular
        .RowAxisLayout xlTabularRow

        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Access the pivot field 'Style'
    Dim pivotField As PivotField
    Set pivotField = pTable.PivotFields("Style")

    ' Set values in pivot worksheet
    pivotWs.Range("A12").Value = "Style field value"
    pivotWs.Range("B12").Value = pivotField.Name

    pivotWs.Range("A14").Value = "New Style field name"
    pivotField.Caption = "My name"
    pivotWs.Range("B14").Value = pivotField.Name

    pivotWs.Range("A15").Value = "Source Style field name"
    pivotWs.Range("B15").Value = pivotField.SourceName
End Sub
```

## OnlyOffice JavaScript Code / Код OnlyOffice JavaScript

```javascript
// Initialize the active worksheet
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

// Create a new worksheet and insert pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Set row layout to Tabular
pivotTable.SetRowAxisLayout("Tabular", false);

// Add data field
pivotTable.AddDataField('Price');

// Access the pivot field 'Style'
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Style');

// Set values in pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Style field value');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetName());

pivotWorksheet.GetRange('A14').SetValue('New Style field name');
pivotField.SetName('My name');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetName());

pivotWorksheet.GetRange('A15').SetValue('Source Style field name');
pivotWorksheet.GetRange('B15').SetValue(pivotField.GetSourceName());
```