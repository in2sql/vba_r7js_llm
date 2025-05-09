# Description
**English:** This code populates an active worksheet with data, creates a pivot table based on that data, adds row and data fields to the pivot table, and sets captions for the data fields.

**Russian:** Этот код заполняет активный лист данными, создает сводную таблицу на основе этих данных, добавляет поля строк и данных в сводную таблицу и устанавливает заголовки для полей данных.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set Region header
oWorksheet.GetRange('C1').SetValue('Style'); // Set Style header
oWorksheet.GetRange('D1').SetValue('Price'); // Set Price header

// Set Region values
oWorksheet.GetRange('B2').SetValue('East'); // Row 2: East
oWorksheet.GetRange('B3').SetValue('West'); // Row 3: West
oWorksheet.GetRange('B4').SetValue('East'); // Row 4: East
oWorksheet.GetRange('B5').SetValue('West'); // Row 5: West

// Set Style values
oWorksheet.GetRange('C2').SetValue('Fancy'); // Row 2: Fancy
oWorksheet.GetRange('C3').SetValue('Fancy'); // Row 3: Fancy
oWorksheet.GetRange('C4').SetValue('Tee');   // Row 4: Tee
oWorksheet.GetRange('C5').SetValue('Tee');   // Row 5: Tee

// Set Price values
oWorksheet.GetRange('D2').SetValue(42.5);    // Row 2: 42.5
oWorksheet.GetRange('D3').SetValue(35.2);    // Row 3: 35.2
oWorksheet.GetRange('D4').SetValue(12.3);    // Row 4: 12.3
oWorksheet.GetRange('D5').SetValue(24.8);    // Row 5: 24.8

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field for 'Sum of Price'
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set captions in the pivot table worksheet
pivotWorksheet.GetRange('A12').SetValue('The Data field caption'); // Caption label
pivotWorksheet.GetRange('B12').SetValue(dataField.GetCaption());     // Data field caption
```

```vba
' Get the active worksheet
Dim ws As Worksheet
Set ws = ThisWorkbook.ActiveSheet

' Set headers
ws.Range("B1").Value = "Region" ' Set Region header
ws.Range("C1").Value = "Style"  ' Set Style header
ws.Range("D1").Value = "Price"  ' Set Price header

' Set Region values
ws.Range("B2").Value = "East"  ' Row 2: East
ws.Range("B3").Value = "West"  ' Row 3: West
ws.Range("B4").Value = "East"  ' Row 4: East
ws.Range("B5").Value = "West"  ' Row 5: West

' Set Style values
ws.Range("C2").Value = "Fancy" ' Row 2: Fancy
ws.Range("C3").Value = "Fancy" ' Row 3: Fancy
ws.Range("C4").Value = "Tee"   ' Row 4: Tee
ws.Range("C5").Value = "Tee"   ' Row 5: Tee

' Set Price values
ws.Range("D2").Value = 42.5    ' Row 2: 42.5
ws.Range("D3").Value = 35.2    ' Row 3: 35.2
ws.Range("D4").Value = 12.3    ' Row 4: 12.3
ws.Range("D5").Value = 24.8    ' Row 5: 24.8

' Define data range
Dim dataRange As Range
Set dataRange = ws.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add

' Create the pivot cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotSheet.Range("A3"), _
    TableName:="PivotTable1")

' Add row fields
pivotTable.PivotFields("Region").Orientation = xlRowField
pivotTable.PivotFields("Style").Orientation = xlRowField

' Add data field
pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

' Set captions in the pivot table worksheet
pivotSheet.Range("A12").Value = "The Data field caption" ' Caption label
pivotSheet.Range("B12").Value = pivotTable.PivotFields("Sum of Price").Caption ' Data field caption
```