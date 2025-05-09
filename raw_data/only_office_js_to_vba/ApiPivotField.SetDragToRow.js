# Create and Manipulate Pivot Table using VBA and OnlyOffice JS  
# Создание и управление сводной таблицей с использованием VBA и OnlyOffice JS

```vba
' VBA code to create and manipulate a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
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
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A1"), TableName:="SalesPivotTable")
    
    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Region").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Modify pivot field properties
    Set pivotField = pivotTable.PivotFields("Region")
    pivotField.Orientation = xlColumnField
    ' VBA does not have a direct "SetDragToRow" equivalent
    
    ' Set additional values
    pivotWs.Range("A13").Value = "Drag to row"
    pivotWs.Range("B13").Value = "False"
    pivotWs.Range("A14").Value = "Try drag Region to rows!"
End Sub
```

```javascript
// JavaScript code to create and manipulate a pivot table using OnlyOffice API

// Get the active worksheet
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

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    rows: ['Style'],          // Add 'Style' as row field
    columns: 'Region',        // Add 'Region' as column field
});

// Add the data field 'Price'
pivotTable.AddDataField('Price');

// Get the active worksheet where pivot table is placed
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field 'Region'
var pivotField = pivotTable.GetPivotFields('Region');

// Set drag to row to false
pivotField.SetDragToRow(false);

// Set additional values
pivotWorksheet.GetRange('A13').SetValue('Drag to row');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToRow());
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to rows!');
```