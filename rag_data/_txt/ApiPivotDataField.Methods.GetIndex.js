**Description / Описание**

This script populates an Excel worksheet with data, creates a pivot table based on that data, and retrieves the index of a specific data field within the pivot table.

Этот скрипт заполняет рабочий лист Excel данными, создает сводную таблицу на основе этих данных и получает индекс конкретного поля данных в сводной таблице.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');   // Region East
oWorksheet.GetRange('B3').SetValue('West');   // Region West
oWorksheet.GetRange('B4').SetValue('East');   // Region East
oWorksheet.GetRange('B5').SetValue('West');   // Region West

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Style Fancy
oWorksheet.GetRange('C3').SetValue('Fancy');  // Style Fancy
oWorksheet.GetRange('C4').SetValue('Tee');    // Style Tee
oWorksheet.GetRange('C5').SetValue('Tee');    // Style Tee

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);     // Price 42.5
oWorksheet.GetRange('D3').SetValue(35.2);     // Price 35.2
oWorksheet.GetRange('D4').SetValue(12.3);     // Price 12.3
oWorksheet.GetRange('D5').SetValue(24.8);     // Price 24.8

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table into a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add rows to pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],               // Add Region and Style as row fields
});

// Add data fields to pivot table
pivotTable.AddDataField('Price');              // Add Price as a data field
pivotTable.AddDataField('Price');              // Add Price again as another data field

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Retrieve the data field 'Sum of Price'
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A15').SetValue('Sum of Price index:'); // Label for index
pivotWorksheet.GetRange('B15').SetValue(dataField.GetIndex());    // Set the index value
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region" ' Set header for Region
oWorksheet.Range("C1").Value = "Style"  ' Set header for Style
oWorksheet.Range("D1").Value = "Price"  ' Set header for Price

' Populate Region data
oWorksheet.Range("B2").Value = "East"   ' Region East
oWorksheet.Range("B3").Value = "West"   ' Region West
oWorksheet.Range("B4").Value = "East"   ' Region East
oWorksheet.Range("B5").Value = "West"   ' Region West

' Populate Style data
oWorksheet.Range("C2").Value = "Fancy"  ' Style Fancy
oWorksheet.Range("C3").Value = "Fancy"  ' Style Fancy
oWorksheet.Range("C4").Value = "Tee"    ' Style Tee
oWorksheet.Range("C5").Value = "Tee"    ' Style Tee

' Populate Price data
oWorksheet.Range("D2").Value = 42.5     ' Price 42.5
oWorksheet.Range("D3").Value = 35.2     ' Price 35.2
oWorksheet.Range("D4").Value = 12.3     ' Price 12.3
oWorksheet.Range("D5").Value = 24.8     ' Price 24.8

' Define data range for pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Insert pivot table into a new worksheet
Dim pivotWorksheet As Worksheet
Set pivotWorksheet = ThisWorkbook.Worksheets.Add
Dim pivotTable As PivotTable
Set pivotTable = pivotWorksheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)

' Add rows to pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField ' Add Region as row field
    .PivotFields("Style").Orientation = xlRowField  ' Add Style as row field
End With

' Add data fields to pivot table
With pivotTable
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Add Price as data field
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Add Price again as data field
End With

' Retrieve the data field 'Sum of Price'
Dim dataField As PivotField
Set dataField = pivotTable.PivotFields("Sum of Price")

' Set values in the pivot worksheet
pivotWorksheet.Range("A15").Value = "Sum of Price index:" ' Label for index
pivotWorksheet.Range("B15").Value = dataField.Position  ' Set the index value
```