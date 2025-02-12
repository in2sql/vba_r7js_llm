## Description / Описание

**English:**  
The code sets up data in the active worksheet, creates a pivot table on a new worksheet, and configures various pivot fields.

**Russian:**  
Код устанавливает данные в активном листе, создает сводную таблицу на новом листе и настраивает различные поля сводной таблицы.

```vba
' VBA Code
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    
    ' Get the active worksheet
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
    
    ' Add a new worksheet for pivot table
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
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Region").Orientation = xlColumnField
        .PivotFields("Price").Orientation = xlDataField
    End With
    
    ' Get the pivot field 'Region'
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set values in pivot worksheet
    pivotWs.Range("A13").Value = "Drag to row"
    pivotWs.Range("B13").Value = (pivotField.Orientation = xlRowField)
End Sub
```

```javascript
// JavaScript Code
// Create pivot table using OnlyOffice API
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

// Insert pivot table on new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add rows and columns fields
pivotTable.AddFields({
    rows: ['Style'],
    columns: 'Region',
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set values in pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to row');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToRow());
```