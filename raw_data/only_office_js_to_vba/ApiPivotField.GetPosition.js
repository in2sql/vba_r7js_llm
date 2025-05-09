**Description:**
This code sets up data in cells B1:D5, creates a pivot table, adds fields, and retrieves the position of the 'Style' field.  
Этот код заполняет данные в ячейках B1:D5, создает сводную таблицу, добавляет поля и получает позицию поля 'Style'.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

// Set Region values
oWorksheet.GetRange('B2').SetValue('East');  // Set Region in B2
oWorksheet.GetRange('B3').SetValue('West');  // Set Region in B3
oWorksheet.GetRange('B4').SetValue('East');  // Set Region in B4
oWorksheet.GetRange('B5').SetValue('West');  // Set Region in B5

// Set Style values
oWorksheet.GetRange('C2').SetValue('Fancy'); // Set Style in C2
oWorksheet.GetRange('C3').SetValue('Fancy'); // Set Style in C3
oWorksheet.GetRange('C4').SetValue('Tee');   // Set Style in C4
oWorksheet.GetRange('C5').SetValue('Tee');   // Set Style in C5

// Set Price values
oWorksheet.GetRange('D2').SetValue(42.5);    // Set Price in D2
oWorksheet.GetRange('D3').SetValue(35.2);    // Set Price in D3
oWorksheet.GetRange('D4').SetValue(12.3);    // Set Price in D4
oWorksheet.GetRange('D5').SetValue(24.8);    // Set Price in D5

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Add data field
pivotTable.AddDataField('Price');

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set values to display the position of 'Style' field
pivotWorksheet.GetRange('A12').SetValue('Style field position'); // Label for position
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetPosition()); // Position value
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region" ' Set header for Region
oWorksheet.Range("C1").Value = "Style"  ' Set header for Style
oWorksheet.Range("D1").Value = "Price"  ' Set header for Price

' Set Region values
oWorksheet.Range("B2").Value = "East"    ' Set Region in B2
oWorksheet.Range("B3").Value = "West"    ' Set Region in B3
oWorksheet.Range("B4").Value = "East"    ' Set Region in B4
oWorksheet.Range("B5").Value = "West"    ' Set Region in B5

' Set Style values
oWorksheet.Range("C2").Value = "Fancy"   ' Set Style in C2
oWorksheet.Range("C3").Value = "Fancy"   ' Set Style in C3
oWorksheet.Range("C4").Value = "Tee"     ' Set Style in C4
oWorksheet.Range("C5").Value = "Tee"     ' Set Style in C5

' Set Price values
oWorksheet.Range("D2").Value = 42.5      ' Set Price in D2
oWorksheet.Range("D3").Value = 35.2      ' Set Price in D3
oWorksheet.Range("D4").Value = 12.3      ' Set Price in D4
oWorksheet.Range("D5").Value = 24.8      ' Set Price in D5

' Define data range for pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotWorksheet As Worksheet
Set pivotWorksheet = Worksheets.Add

' Create the pivot cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A3"), TableName:="PivotTable1")

' Add row fields
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add data field
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
End With

' Get the position of the 'Style' field
Dim pivotField As PivotField
Set pivotField = pivotTable.PivotFields("Style")

' Set values to display the position of 'Style' field
pivotWorksheet.Range("A12").Value = "Style field position" ' Label for position
pivotWorksheet.Range("B12").Value = pivotField.Position      ' Position value
```