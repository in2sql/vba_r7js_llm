**Description / Описание**

This script initializes data in an Excel worksheet, creates a pivot table based on that data, and retrieves specific information from the pivot table.

Этот скрипт инициализирует данные в листе Excel, создает сводную таблицу на основе этих данных и извлекает определенную информацию из сводной таблицы.

---

```vba
' VBA Code to initialize data, create a pivot table, and retrieve pivot item information

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    Dim pivotItem As PivotItem
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Initialize headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Initialize data
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
        TableName:="PivotTable1")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlPageField
        .PivotFields("Region").Orientation = xlRowField
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount
    End With
    
    ' Get the pivot field and item
    Set pivotField = pivotTable.PivotFields("Style")
    Set pivotItem = pivotField.PivotItems(1)
    
    ' Set values based on pivot item
    pivotWs.Range("A15").Value = pivotItem.Name & " parent:"
    pivotWs.Range("B15").Value = pivotItem.Parent.Name
End Sub
```

```javascript
// JavaScript Code to initialize data, create a pivot table, and retrieve pivot item information using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Initialize headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Initialize data
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

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    pages: ['Style'],
    rows: 'Region',
});

// Add data field to the pivot table
pivotTable.AddDataField('Style');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Get the first pivot item
var pivotItem = pivotField.GetPivotItems()[0];

// Set values based on pivot item
pivotWorksheet.GetRange('A15').SetValue(pivotItem.GetName() + ' parent:');
pivotWorksheet.GetRange('B15').SetValue(pivotItem.GetParent().GetName());
```