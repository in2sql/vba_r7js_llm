# Create and Populate a Worksheet with Data and Pivot Table  
# Создание и заполнение рабочего листа данными и сводной таблицей

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set Region values
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set Style values
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price values
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add rows and columns to the pivot table
pivotTable.AddFields({
	rows: ['Style'],
	columns: 'Region',
});

// Add the Price field as a data field
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set additional values in the pivot table sheet
pivotWorksheet.GetRange('A13').SetValue('Drag to page');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToPage());
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Set Region values
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Set Style values
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Set Price values
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range
Dim dataRange As Range
Set dataRange = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotWorksheet As Worksheet
Set pivotWorksheet = ThisWorkbook.Worksheets.Add
pivotWorksheet.Name = "PivotTableSheet"

' Create the pivot table cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")

' Add rows and columns to the pivot table
With pivotTable
    .PivotFields("Style").Orientation = xlRowField
    .PivotFields("Region").Orientation = xlColumnField
    .PivotFields("Price").Orientation = xlDataField
    .PivotFields("Price").Function = xlSum
End With

' Set additional values in the pivot table sheet
pivotWorksheet.Range("A13").Value = "Drag to page"
pivotWorksheet.Range("B13").Value = "Region" ' VBA does not have a direct equivalent of GetDragToPage
```