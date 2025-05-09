# Code Description / Описание кода

This script populates an Excel worksheet with data, creates a pivot table based on that data, and displays certain pivot table field properties.

Этот скрипт заполняет рабочий лист Excel данными, создаёт сводную таблицу на основе этих данных и отображает определённые свойства полей сводной таблицы.

```vba
' VBA Code Equivalent

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    
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
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set Price data
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
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add 'Region' to row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Region").Position = 1
    
    ' Add 'Price' to data fields
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Display 'Showing in Axis' properties
    pivotWs.Range("A12").Value = "Region showing in axis"
    pivotWs.Range("B12").Value = pivotTable.PivotFields("Region").Orientation = xlRowField
    
    pivotWs.Range("A13").Value = "Style showing in axis"
    pivotWs.Range("B13").Value = pivotTable.PivotFields("Style").Orientation <> xlHidden
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
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
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Region' as row field
pivotTable.AddFields({
    rows: ['Region'],
});

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Set and display 'Region showing in axis'
pivotWorksheet.GetRange('A12').SetValue('Region showing in axis');
pivotWorksheet.GetRange('B12').SetValue(pivotTable.GetPivotFields('Region').GetShowingInAxis());

// Set and display 'Style showing in axis'
pivotWorksheet.GetRange('A13').SetValue('Style showing in axis');
pivotWorksheet.GetRange('B13').SetValue(pivotTable.GetPivotFields('Style').GetShowingInAxis());
```