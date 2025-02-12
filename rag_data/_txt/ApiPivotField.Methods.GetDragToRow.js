---

**Description:**

English:  
This script populates an active worksheet with data, creates a pivot table on a new worksheet, configures the pivot fields for rows and columns, adds a data field, and updates specific cells with pivot field information.

Russian:  
Этот скрипт заполняет активный лист данными, создает сводную таблицу на новом листе, настраивает поля сводной таблицы для строк и столбцов, добавляет поле данных и обновляет определенные ячейки информацией о поле сводной таблицы.

---

**VBA Code:**

```vba
' VBA code to replicate the OnlyOffice JS functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
    ws.Range("B1").Value = "Region"    ' Set header for Region
    ws.Range("C1").Value = "Style"     ' Set header for Style
    ws.Range("D1").Value = "Price"     ' Set header for Price
    
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
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add 'Style' to Rows
    pivotTable.PivotFields("Style").Orientation = xlRowField
    pivotTable.PivotFields("Style").Position = 1
    
    ' Add 'Region' to Columns
    pivotTable.PivotFields("Region").Orientation = xlColumnField
    pivotTable.PivotFields("Region").Position = 1
    
    ' Add 'Price' to Values
    pivotTable.PivotFields("Price").Orientation = xlDataField
    pivotTable.PivotFields("Price").Function = xlSum
    pivotTable.PivotFields("Price").Name = "Sum of Price"
    
    ' Update specific cells with information
    pivotWs.Range("A13").Value = "Drag to row"  ' Set label
    Set pivotField = pivotTable.PivotFields("Region")
    pivotWs.Range("B13").Value = pivotField.Orientation = xlRowField ' Set drag to row status
    
End Sub
```

---

**OnlyOffice JS Code:**

```javascript
// OnlyOffice JS code to populate data, create pivot table, and update cells

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

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

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Configure pivot table fields
pivotTable.AddFields({
    rows: ['Style'],       // Add 'Style' to Rows
    columns: 'Region'      // Add 'Region' to Columns
});

// Add 'Price' as data field
pivotTable.AddDataField('Price'); // Add 'Price' to Values

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Update specific cells with pivot field information
pivotWorksheet.GetRange('A13').SetValue('Drag to row'); // Set label
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToRow()); // Set drag to row status
```