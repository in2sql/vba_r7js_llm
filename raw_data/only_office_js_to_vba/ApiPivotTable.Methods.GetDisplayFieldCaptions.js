# Description / Описание

**English:** This script populates data into a worksheet, creates a pivot table from the data, and adds fields to the pivot table. It also sets values in another worksheet based on the pivot table.

**Русский:** Этот скрипт заполняет данные в листе, создает сводную таблицу из данных и добавляет поля в сводную таблицу. Он также устанавливает значения в другом листе на основе сводной таблицы.

```vba
' VBA code equivalent to the OnlyOffice JS API example

Sub CreatePivotTable()
    ' Set reference to the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Populate headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Populate Region data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Populate Style data
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Populate Price data
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A3"), TableName:="PivotTable1")
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Add row and column fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlColumnField
    
    ' Set values in the pivot worksheet
    pivotWorksheet.Range("A9").Value = "Display field captions"
    pivotWorksheet.Range("B9").Value = "Sum of Price" ' Example caption
End Sub
```

```javascript
// OnlyOffice JS code equivalent to the VBA example

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate headers
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add data field
pivotTable.AddDataField('Price');

// Add row and column fields
pivotTable.AddFields({
    rows: 'Region',
    columns: 'Style',
});

// Get the active worksheet for the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A9').SetValue('Display field captions');
pivotWorksheet.GetRange('B9').SetValue(pivotTable.GetDisplayFieldCaptions());
```