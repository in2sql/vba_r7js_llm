### Description / Описание
**English:** This script sets up a worksheet with data and creates a pivot table.

**Russian:** Этот скрипт настраивает рабочий лист с данными и создаёт сводную таблицу.

---

### Excel VBA Code
```vba
' This VBA script sets up data in Sheet1 and creates a pivot table on a new worksheet.

Sub CreatePivotTable()
    ' Define variables
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
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
    
    ' Define data range
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set dataRange = ws.Range(ws.Cells(1, 2), ws.Cells(lastRow, lastCol))
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Sheets.Add(After:=ws)
    pivotWs.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add Row Fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .RowAxisLayout xlTabularRow
    End With
    
    ' Add Data Field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set values in pivot worksheet
    pivotWs.Range("A12").Value = "Style field value"
    pivotWs.Range("B12").Value = pivotTable.PivotFields("Style").CurrentPage
    
    pivotWs.Range("A14").Value = "New Style field value"
    pivotTable.PivotFields("Style").CurrentPage = "My value"
    pivotWs.Range("B14").Value = pivotTable.PivotFields("Style").CurrentPage
End Sub
```

---

### OnlyOffice JavaScript Code
```javascript
// This JavaScript code sets up data in the active sheet and creates a pivot table on a new worksheet.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate data in column B
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate data in column C
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate data in column D
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Set row axis layout to tabular form
pivotTable.SetRowAxisLayout("Tabular", false);

// Add 'Price' as a data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' field from the pivot table
var pivotField = pivotTable.GetPivotFields('Style');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Style field value');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetValue());

pivotWorksheet.GetRange('A14').SetValue('New Style field value');
pivotField.SetValue('My value');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetValue());
```