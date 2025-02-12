**Description / Описание:**

This code populates an Excel sheet with data, creates a pivot table, modifies a data field name, and updates specific cells with information about the data field.

Этот код заполняет лист Excel данными, создает сводную таблицу, изменяет имя поля данных и обновляет определенные ячейки информацией о поле данных.

```vba
' VBA Code

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim ws As Worksheet
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
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the Pivot Table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create the Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A1"), TableName:="MyPivotTable")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the data field
    Dim dataField As PivotField
    Set dataField = pivotTable.PivotFields("Sum of Price")
    
    ' Set cell A12 and B12
    pivotWs.Range("A12").Value = "Data field name"
    pivotWs.Range("B12").Value = dataField.Name
    
    ' Rename the data field
    dataField.Name = "My Sum of Price"
    
    ' Set cell A13 and B13
    pivotWs.Range("A13").Value = "New Data field name"
    pivotWs.Range("B13").Value = dataField.Name
End Sub
```

```javascript
// OnlyOffice JS Code

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
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new Pivot Table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Sum of Price' data field
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set cell A12 and B12
pivotWorksheet.GetRange('A12').SetValue('Data field name');
pivotWorksheet.GetRange('B12').SetValue(dataField.GetName());

// Rename the data field
dataField.SetName('My Sum of Price');

// Set cell A13 and B13
pivotWorksheet.GetRange('A13').SetValue('New Data field name');
pivotWorksheet.GetRange('B13').SetValue(dataField.GetName());
```