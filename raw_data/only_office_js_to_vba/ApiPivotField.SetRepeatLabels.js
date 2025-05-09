# Description / Описание

This code sets up data in an Excel worksheet, creates a pivot table based on that data, configures the pivot table fields, and updates specific cells to reflect the pivot table's settings.

Этот код заполняет данные в рабочем листе Excel, создает сводную таблицу на основе этих данных, настраивает поля сводной таблицы и обновляет определенные ячейки, чтобы отразить настройки сводной таблицы.

## Excel VBA Code

```vba
' This VBA code sets up data in a worksheet, creates a pivot table, configures it,
' and updates specific cells based on the pivot table's settings.

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    
    ' Set the active worksheet
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
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Set pivotTable = pivotWs.PivotTables.Add(PivotCache:=ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=dataRange), TableDestination:=pivotWs.Range("A3"), TableName:="PivotTable1")
    
    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
        .RowAxisLayout xlTabularRow
    End With
    
    ' Configure the 'Region' field to repeat labels
    Set pivotField = pivotTable.PivotFields("Region")
    pivotField.RepeatLabels = True
    
    ' Update specific cells with the repeat labels information
    pivotWs.Range("A12").Value = "Region repeat labels"
    pivotWs.Range("B12").Value = pivotField.RepeatLabels
End Sub
```

## OnlyOffice JavaScript Code

```javascript
// This JavaScript code sets up data in an OnlyOffice sheet, creates a pivot table,
// configures its fields, and updates specific cells based on the pivot table's settings.

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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add 'Price' as a data field
pivotTable.AddDataField('Price');

// Set the row axis layout to 'Tabular'
pivotTable.SetRowAxisLayout('Tabular');

// Configure the 'Region' field to repeat labels
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.SetRepeatLabels(true);

// Update specific cells with repeat labels information
pivotWorksheet.GetRange('A12').SetValue('Region repeat labels');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetRepeatLabels());
```