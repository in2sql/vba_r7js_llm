# Description
**English:** The script populates specific cells with data, creates a pivot table based on the data range, adds row and column fields, and defines data fields in the pivot table.

**Russian:** Скрипт заполняет определённые ячейки данными, создаёт сводную таблицу на основе диапазона данных, добавляет поля строк и столбцов, а также определяет поля данных в сводной таблице.

```vba
' VBA code to populate cells and create a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
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
    
    ' Add a new worksheet for pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create pivot table
    Set pt = ptCache.CreatePivotTable(TableDestination:=pivotWs.Range("A1"), TableName:="SalesPivotTable")
    
    ' Add fields to pivot table
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Add labels
    pivotWs.Range("A13").Value = "Drag to column"
    pivotWs.Range("B13").Value = pt.PivotFields("Region").Orientation = xlColumnField
End Sub
```

```javascript
// OnlyOffice JavaScript code to populate cells and create a pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set cell B1 to 'Region'
oWorksheet.GetRange('C1').SetValue('Style'); // Set cell C1 to 'Style'
oWorksheet.GetRange('D1').SetValue('Price'); // Set cell D1 to 'Price'

// Set data
oWorksheet.GetRange('B2').SetValue('East'); // Set cell B2 to 'East'
oWorksheet.GetRange('B3').SetValue('West'); // Set cell B3 to 'West'
oWorksheet.GetRange('B4').SetValue('East'); // Set cell B4 to 'East'
oWorksheet.GetRange('B5').SetValue('West'); // Set cell B5 to 'West'

oWorksheet.GetRange('C2').SetValue('Fancy'); // Set cell C2 to 'Fancy'
oWorksheet.GetRange('C3').SetValue('Fancy'); // Set cell C3 to 'Fancy'
oWorksheet.GetRange('C4').SetValue('Tee');   // Set cell C4 to 'Tee'
oWorksheet.GetRange('C5').SetValue('Tee');   // Set cell C5 to 'Tee'

oWorksheet.GetRange('D2').SetValue(42.5);    // Set cell D2 to 42.5
oWorksheet.GetRange('D3').SetValue(35.2);    // Set cell D3 to 35.2
oWorksheet.GetRange('D4').SetValue(12.3);    // Set cell D4 to 12.3
oWorksheet.GetRange('D5').SetValue(24.8);    // Set cell D5 to 24.8

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    columns: ['Style'], // Set 'Style' as column field
    rows: 'Region',     // Set 'Region' as row field
});

// Add data field to pivot table
pivotTable.AddDataField('Price'); // Add 'Price' as data field

// Get pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get pivot field for 'Region'
var pivotField = pivotTable.GetPivotFields('Region');

// Set labels
pivotWorksheet.GetRange('A13').SetValue('Drag to column'); // Set cell A13
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToColumn()); // Set cell B13
```