# Description

This code sets up a worksheet with region, style, and price data, then creates a pivot table based on that data, organizing it by style and region, and adding price as a data field.

Этот код настраивает лист с данными региона, стиля и цены, затем создает сводную таблицу на основе этих данных, организуя ее по стилю и региону и добавляя цену как поле данных.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set value for cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set value for cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set value for cell D1

// Set Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set value for cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set value for cell B3
oWorksheet.GetRange('B4').SetValue('East');   // Set value for cell B4
oWorksheet.GetRange('B5').SetValue('West');   // Set value for cell B5

// Set Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set value for cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set value for cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set value for cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set value for cell C5

// Set Price data
oWorksheet.GetRange('D2').SetValue(42.5);      // Set value for cell D2
oWorksheet.GetRange('D3').SetValue(35.2);      // Set value for cell D3
oWorksheet.GetRange('D4').SetValue(12.3);      // Set value for cell D4
oWorksheet.GetRange('D5').SetValue(24.8);      // Set value for cell D5

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to the pivot table
pivotTable.AddFields({
    rows: ['Style'],         // Add 'Style' as row field
    columns: 'Region',      // Add 'Region' as column field
});

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Get the newly created pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' field in the pivot table
var pivotField = pivotTable.GetPivotFields('Region');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to data');                // Set value for cell A13
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToData());     // Set value for cell B13
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region"  ' Set value for cell B1
oWorksheet.Range("C1").Value = "Style"   ' Set value for cell C1
oWorksheet.Range("D1").Value = "Price"   ' Set value for cell D1

' Set Region data
oWorksheet.Range("B2").Value = "East"     ' Set value for cell B2
oWorksheet.Range("B3").Value = "West"     ' Set value for cell B3
oWorksheet.Range("B4").Value = "East"     ' Set value for cell B4
oWorksheet.Range("B5").Value = "West"     ' Set value for cell B5

' Set Style data
oWorksheet.Range("C2").Value = "Fancy"    ' Set value for cell C2
oWorksheet.Range("C3").Value = "Fancy"    ' Set value for cell C3
oWorksheet.Range("C4").Value = "Tee"      ' Set value for cell C4
oWorksheet.Range("C5").Value = "Tee"      ' Set value for cell C5

' Set Price data
oWorksheet.Range("D2").Value = 42.5       ' Set value for cell D2
oWorksheet.Range("D3").Value = 35.2       ' Set value for cell D3
oWorksheet.Range("D4").Value = 12.3       ' Set value for cell D4
oWorksheet.Range("D5").Value = 24.8       ' Set value for cell D5

' Define data range for pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Insert a new pivot table in a new worksheet
Dim pivotWorksheet As Worksheet
Dim pivotTable As PivotTable
Set pivotWorksheet = Worksheets.Add
Set pivotTable = pivotWorksheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)

' Add row and column fields to the pivot table
With pivotTable
    .PivotFields("Style").Orientation = xlRowField     ' Add 'Style' as row field
    .PivotFields("Region").Orientation = xlColumnField ' Add 'Region' as column field
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Add 'Price' as data field
End With

' Set values in the pivot worksheet
pivotWorksheet.Range("A13").Value = "Drag to data" ' Set value for cell A13
pivotWorksheet.Range("B13").Value = "Drag to data not directly applicable" ' VBA does not have a direct equivalent to GetDragToData()
```