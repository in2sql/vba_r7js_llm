### Description
This script sets up data in an OnlyOffice spreadsheet, creates a pivot table based on that data, and modifies the pivot table by removing a field after a delay.

Этот скрипт заполняет данные в таблице OnlyOffice, создает сводную таблицу на основе этих данных и изменяет сводную таблицу, удаляя поле после задержки.

### OnlyOffice JS Code
```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data for Region
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set data for Style
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set data for Price
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to the pivot table
pivotTable.AddFields({
	rows: 'Region',
	columns: 'Style',
});

// Get the active sheet where the pivot table is located
var pivotWorksheet = Api.GetActiveSheet();

// Add a data field to the pivot table
pivotTable.AddDataField('Price');

// Get the 'Region' field from the pivot table
var pivotField = pivotTable.GetPivotFields('Region');

// Set a message in cell A10
pivotWorksheet.GetRange('A10').SetValue('The Region field will be removed soon');

// Remove the 'Region' field after a 5-second delay
setTimeout(function () {
	pivotField.Remove();
}, 5000);
```

### Excel VBA Code
```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Set data for Region
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Set data for Style
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Set data for Price
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add
pivotSheet.Name = "PivotSheet"

' Create the pivot cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A3"), TableName:="PivotTable1")

' Add row and column fields to the pivot table
pivotTable.PivotFields("Region").Orientation = xlRowField
pivotTable.PivotFields("Style").Orientation = xlColumnField

' Add data field to the pivot table
pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

' Set a message in cell A10
pivotSheet.Range("A10").Value = "The Region field will be removed soon"

' Remove the 'Region' field after a 5-second delay
Application.OnTime Now + TimeValue("00:00:05"), "RemoveRegionField"

' Subroutine to remove the 'Region' field
Sub RemoveRegionField()
    On Error Resume Next
    pivotTable.PivotFields("Region").Orientation = xlHidden
End Sub
```