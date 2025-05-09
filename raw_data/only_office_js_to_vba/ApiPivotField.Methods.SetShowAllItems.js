## Description / Описание

This script sets up data in the active worksheet, creates a pivot table in a new worksheet, configures the pivot table fields, and sets some values based on pivot table properties.

Этот скрипт устанавливает данные в активном листе, создает сводную таблицу на новом листе, настраивает поля сводной таблицы и устанавливает некоторые значения на основе свойств сводной таблицы.

```vba
' VBA Code to replicate OnlyOffice JS functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Tee"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set Price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the Pivot Table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add Row Fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add Data Field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the Pivot Field 'Style'
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Set ShowAllItems to True
    pivotField.ShowAllItems = True
    
    ' Set values based on Pivot Field properties
    pivotWs.Range("A12").Value = "Style get show all items"
    pivotWs.Range("B12").Value = pivotField.ShowAllItems
End Sub
```

```javascript
// OnlyOffice JS Code to set up data and create a pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set Region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set Style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Tee');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set ShowAllItems to true for 'Style' field
pivotField.SetShowAllItems(true);

// Set values based on Pivot Field properties
pivotWorksheet.GetRange('A12').SetValue('Style get show all items');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetShowAllItems());
```