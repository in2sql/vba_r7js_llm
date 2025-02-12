# Description / Описание

This code creates and populates an Excel worksheet with region, style, and price data, inserts a pivot table based on this data, adds specific fields to the pivot table, and then retrieves and displays the captions of the pivot table's style items.

Этот код создает и заполняет рабочий лист Excel данными о регионе, стиле и цене, вставляет на основе этих данных сводную таблицу, добавляет определенные поля в сводную таблицу, а затем извлекает и отображает подписи элементов стиля сводной таблицы.

## VBA Code

```vba
' VBA code to replicate the OnlyOffice JavaScript functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
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
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlColumnField
        .PivotFields("Region").Orientation = xlRowField
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount
    End With
    
    ' Get pivot items for 'Style' field
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Write header for style item captions
    pivotWs.Cells(15, 1).Value = "Style item captions"
    
    ' Loop through pivot items and write captions
    For i = 1 To pivotField.PivotItems.Count
        pivotWs.Cells(15 + i - 1, 2).Value = pivotField.PivotItems(i).Name
    Next i
End Sub
```

## JavaScript Code

```javascript
// JavaScript code using OnlyOffice API to create and manipulate a pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    columns: ['Style'], // Add 'Style' as column field
    rows: 'Region',     // Add 'Region' as row field
});

// Add 'Style' as data field with count aggregation
pivotTable.AddDataField('Style');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Get all pivot items for 'Style' field
var pivotItems = pivotField.GetPivotItems();

// Set header for style item captions
pivotWorksheet.GetRangeByNumber(15, 0).SetValue('Style item captions');

// Loop through pivot items and set their captions
for (var i = 0; i < pivotItems.length; i += 1) {
    pivotWorksheet.GetRangeByNumber(15 + i, 1).SetValue(pivotItems[i].GetCaption());
}
```