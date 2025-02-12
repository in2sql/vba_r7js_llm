**Description / Описание**

This script sets up data in a worksheet, creates a pivot table from that data, and retrieves a property from the pivot table's field.

Этот скрипт устанавливает данные в листе, создает сводную таблицу на основе этих данных и получает свойство поля сводной таблицы.

```javascript
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
oWorksheet.GetRange('C3').SetValue('Tee');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Region and Style as row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add Price as a data field
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set description in cell A12
pivotWorksheet.GetRange('A12').SetValue('Style get show all items');

// Set the value of ShowAllItems property in cell B12
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetShowAllItems());
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set headers
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
oWorksheet.Range("C3").Value = "Tee"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Populate Price data
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Create a new worksheet for the pivot table
Dim pivotWorksheet As Worksheet
Set pivotWorksheet = Worksheets.Add

' Create the pivot cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRef)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotWorksheet.Range("A3"), _
    TableName:="PivotTable1")

' Add Region and Style as row fields
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add Price as a data field
    .PivotFields("Price").Orientation = xlDataField
    .PivotFields("Price").Function = xlSum
End With

' Get the 'Style' pivot field
Dim pivotFieldStyle As PivotField
Set pivotFieldStyle = pivotTable.PivotFields("Style")

' Set description in cell A12
pivotWorksheet.Range("A12").Value = "Style get show all items"

' Set the value of ShowAllItems property in cell B12
pivotWorksheet.Range("B12").Value = pivotFieldStyle.ShowAllItems
```