**Description / Описание**

This script populates an Excel worksheet with region, style, and price data, creates a pivot table based on this data, adds specific fields to the pivot table, and retrieves the position of the "Style" field within the pivot table.

Этот скрипт заполняет рабочий лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, добавляет определенные поля в сводную таблицу и получает позицию поля "Style" в сводной таблице.

```javascript
// JavaScript Code using OnlyOffice API

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

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add rows fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Add Price as a data field
pivotTable.AddDataField('Price');

// Get the pivot field for 'Style'
var pivotField = pivotTable.GetPivotFields('Style');

// Set the position of the 'Style' field in cells A12 and B12
pivotWorksheet.GetRange('A12').SetValue('Style field position');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetPosition());
```

```vba
' VBA Code equivalent to the OnlyOffice API script

Sub CreatePivotTable()
    ' Declare variables
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
    
    ' Populate Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Populate Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Populate Price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add Row Fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Region").Position = 1
    pivotTable.PivotFields("Style").Orientation = xlRowField
    pivotTable.PivotFields("Style").Position = 2
    
    ' Add Data Field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the pivot field for 'Style' and its position
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Set the position value in cells A12 and B12
    pivotWs.Range("A12").Value = "Style field position"
    pivotWs.Range("B12").Value = pivotField.Position
End Sub
```