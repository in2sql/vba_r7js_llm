# Description / Описание

**English:**  
This script populates an Excel worksheet with sample data, creates a pivot table on a new worksheet, adds fields to the pivot table, inserts a note in the pivot worksheet, and removes a pivot field after a delay.

**Русский:**  
Этот скрипт заполняет лист Excel примерными данными, создает сводную таблицу на новом листе, добавляет поля в сводную таблицу, вставляет примечание на лист со сводной таблицей и удаляет одно из полей сводной таблицы после задержки.

---

```vba
' VBA code to populate data, create pivot table, manipulate fields, and remove a field after delay

Sub CreatePivotTable()
    ' Declare variables
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim pivotWs As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    
    ' Set active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
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
    
    ' Define data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A1"), TableName:="MyPivotTable")
    
    ' Add fields to pivot table
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlColumnField
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Add note to pivot worksheet
    pivotWs.Range("A10").Value = "The Region field will be removed soon"
    
    ' Schedule removal of 'Region' field after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "RemoveRegionField"
End Sub

Sub RemoveRegionField()
    ' Declare variables
    Dim pivotTbl As PivotTable
    Dim pivotFld As PivotField
    
    ' Set pivot table (assumes only one pivot table in the workbook)
    Set pivotTbl = ThisWorkbook.PivotTables("MyPivotTable")
    
    ' Set pivot field
    Set pivotFld = pivotTbl.PivotFields("Region")
    
    ' Remove the pivot field
    pivotFld.Orientation = xlHidden
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to populate data, create pivot table, manipulate fields, and remove a field after delay

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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to pivot table
pivotTable.AddFields({
	rows: 'Region',
	columns: 'Style',
});

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Add data field to pivot table
pivotTable.AddDataField('Price');

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Insert a note in the pivot worksheet
pivotWorksheet.GetRange('A10').SetValue('The Region field will be removed soon');

// Remove the 'Region' field after 5 seconds
setTimeout(function () {
	pivotField.Remove();
}, 5000);
```