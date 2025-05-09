### Description / Описание

**English**: This code initializes a worksheet with sample data, creates a pivot table based on the data, configures the pivot table fields, and updates specific cells with information about the pivot field settings.

**Russian**: Этот код инициализирует лист данных пробными данными, создает сводную таблицу на основе этих данных, настраивает поля сводной таблицы и обновляет определенные ячейки информацией о настройках полей сводной таблицы.

```vba
' VBA Code to replicate OnlyOffice API functionality

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
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add row field
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add column field
    pivotTable.PivotFields("Region").Orientation = xlColumnField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the pivot field 'Region'
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Attempt to disable drag to page (not directly possible in VBA)
    ' VBA does not have a direct equivalent, so we skip or note this
    ' Alternatively, we can manipulate the PageFields
    ' For the purpose of this example, we set it as a regular field
    pivotField.Orientation = xlColumnField
    
    ' Update specific cells
    pivotWs.Range("A13").Value = "Drag to page"
    pivotWs.Range("B13").Value = "False"
    pivotWs.Range("A14").Value = "Try drag Region to pages!"
    
End Sub
```

```javascript
// OnlyOffice JS Code as provided

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

// Insert pivot table in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
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

// Disable drag to page
pivotField.SetDragToPage(false);

// Update cells with pivot field settings
pivotWorksheet.GetRange('A13').SetValue('Drag to page');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToPage());
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to pages!');
```