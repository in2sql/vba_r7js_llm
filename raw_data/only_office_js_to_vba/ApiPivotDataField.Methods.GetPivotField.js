### Description / Описание
This script sets up a worksheet with region, style, and price data, creates a pivot table based on this data, and retrieves specific information about the pivot table fields.

Этот скрипт настраивает лист с данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных и извлекает конкретную информацию о полях сводной таблицы.

```vba
' VBA Code to set up worksheet and create pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataField As PivotField
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data values
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
    
    ' Create a new worksheet for the pivot table
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
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add data fields
    With pivotTable
        .PivotFields("Price").Orientation = xlDataField
        .PivotFields("Price").Function = xlSum
        .PivotFields("Price").Name = "Sum of Price"
        .PivotFields("Price").Orientation = xlDataField
    End With
    
    ' Get the data field
    Set dataField = pivotTable.PivotFields("Sum of Price")
    
    ' Set values in pivot worksheet
    pivotWs.Range("A15").Value = "Sum of Price position:"
    pivotWs.Range("B15").Value = dataField.Position
    
    pivotWs.Range("A16").Value = "Price position:"
    pivotWs.Range("B16").Value = dataField.DataRange.Column
End Sub
```

```javascript
// OnlyOffice JS Code to set up worksheet and create pivot table

var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Set data values
oWorksheet.GetRange('B2').SetValue('East');    // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');    // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');    // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');    // Set 'West' in cell B5

oWorksheet.GetRange('C2').SetValue('Fancy');   // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');   // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');     // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');     // Set 'Tee' in cell C5

oWorksheet.GetRange('D2').SetValue(42.5);      // Set 42.5 in cell D2
oWorksheet.GetRange('D3').SetValue(35.2);      // Set 35.2 in cell D3
oWorksheet.GetRange('D4').SetValue(12.3);      // Set 12.3 in cell D4
oWorksheet.GetRange('D5').SetValue(24.8);      // Set 24.8 in cell D5

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Add 'Region' and 'Style' as row fields
});

// Add data fields
pivotTable.AddDataField('Price');    // Add 'Price' as data field
pivotTable.AddDataField('Price');    // Add 'Price' again as data field

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Sum of Price' data field
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in pivot worksheet
pivotWorksheet.GetRange('A15').SetValue('Sum of Price position:'); // Set label in A15
pivotWorksheet.GetRange('B15').SetValue(dataField.GetIndex());        // Set position in B15

pivotWorksheet.GetRange('A16').SetValue('Price position:');           // Set label in A16
pivotWorksheet.GetRange('B16').SetValue(dataField.GetPivotField().GetIndex()); // Set position in B16
```