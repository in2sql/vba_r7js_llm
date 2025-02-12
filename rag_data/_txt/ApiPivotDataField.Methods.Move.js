**Description / Описание**

This code populates data into an Excel worksheet, creates a pivot table, adds fields and data fields, and modifies the pivot table after a delay.

Этот код заполняет данные в рабочем листе Excel, создает сводную таблицу, добавляет поля и поля данных, а также изменяет сводную таблицу после задержки.

---

**VBA Code**

```vba
' VBA code to populate data, create a pivot table, and modify it after a delay

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add Row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add Data fields
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price 2", xlSum
    
    ' Set a cell value
    pivotWs.Range("A16").Value = "Sum of Price will be moved soon"
    
    ' Use Application.OnTime to delay action
    Application.OnTime Now + TimeValue("00:00:05"), "MoveDataField"
End Sub

' Subroutine to move data field to rows after delay
Sub MoveDataField()
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets("PivotTableSheet")
    
    Dim pivotTable As PivotTable
    Set pivotTable = pivotWs.PivotTables("PivotTable1")
    
    Dim dataField As PivotField
    Set dataField = pivotTable.PivotFields("Sum of Price")
    
    ' Move data field to Row
    dataField.Orientation = xlRowField
    dataField.Position = 1
End Sub
```

---

**OnlyOffice JavaScript Code**

```javascript
// JavaScript code to populate data, create a pivot table, and modify it after a delay

var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add Data fields
pivotTable.AddDataField('Price');
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get data fields
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set a cell value
pivotWorksheet.GetRange('A16').SetValue('Sum of Price will be moved soon');

// After 5 seconds, move the data field to rows
setTimeout(function() {
    dataField.Move("Rows");
}, 5000);
```