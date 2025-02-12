**Description / Описание**

*English:*  
This script populates an active worksheet with regional sales data, creates a pivot table summarizing the data by Region and Style, and updates the pivot table's data field label.

*Russian:*  
Этот скрипт заполняет активный лист данными о продажах по регионам, создает сводную таблицу, обобщающую данные по Региону и Стилю, и обновляет название поля данных в сводной таблице.

---

### JavaScript Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set cell B1 to 'Region'
oWorksheet.GetRange('C1').SetValue('Style');  // Set cell C1 to 'Style'
oWorksheet.GetRange('D1').SetValue('Price');  // Set cell D1 to 'Price'

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set cell B2 to 'East'
oWorksheet.GetRange('B3').SetValue('West');   // Set cell B3 to 'West'
oWorksheet.GetRange('B4').SetValue('East');   // Set cell B4 to 'East'
oWorksheet.GetRange('B5').SetValue('West');   // Set cell B5 to 'West'

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set cell C2 to 'Fancy'
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set cell C3 to 'Fancy'
oWorksheet.GetRange('C4').SetValue('Tee');    // Set cell C4 to 'Tee'
oWorksheet.GetRange('C5').SetValue('Tee');    // Set cell C5 to 'Tee'

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);      // Set cell D2 to 42.5
oWorksheet.GetRange('D3').SetValue(35.2);      // Set cell D3 to 35.2
oWorksheet.GetRange('D4').SetValue(12.3);      // Set cell D4 to 12.3
oWorksheet.GetRange('D5').SetValue(24.8);      // Set cell D5 to 24.8

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field to pivot table
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field from the pivot table
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set label and value in pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Data field value');       // Set cell A12
pivotWorksheet.GetRange('B12').SetValue(dataField.GetValue());     // Set cell B12 with data field value

// Update the data field label
dataField.SetValue('My Sum of Price');                            // Rename data field
pivotWorksheet.GetRange('A13').SetValue('New Data field value');   // Set cell A13
pivotWorksheet.GetRange('B13').SetValue(dataField.GetValue());     // Set cell B13 with updated data field value
```

### VBA Code

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region" ' Set cell B1 to 'Region'
oWorksheet.Range("C1").Value = "Style"  ' Set cell C1 to 'Style'
oWorksheet.Range("D1").Value = "Price"  ' Set cell D1 to 'Price'

' Populate Region data
oWorksheet.Range("B2").Value = "East"   ' Set cell B2 to 'East'
oWorksheet.Range("B3").Value = "West"   ' Set cell B3 to 'West'
oWorksheet.Range("B4").Value = "East"   ' Set cell B4 to 'East'
oWorksheet.Range("B5").Value = "West"   ' Set cell B5 to 'West'

' Populate Style data
oWorksheet.Range("C2").Value = "Fancy"  ' Set cell C2 to 'Fancy'
oWorksheet.Range("C3").Value = "Fancy"  ' Set cell C3 to 'Fancy'
oWorksheet.Range("C4").Value = "Tee"    ' Set cell C4 to 'Tee'
oWorksheet.Range("C5").Value = "Tee"    ' Set cell C5 to 'Tee'

' Populate Price data
oWorksheet.Range("D2").Value = 42.5      ' Set cell D2 to 42.5
oWorksheet.Range("D3").Value = 35.2      ' Set cell D3 to 35.2
oWorksheet.Range("D4").Value = 12.3      ' Set cell D4 to 12.3
oWorksheet.Range("D5").Value = 24.8      ' Set cell D5 to 24.8

' Define data range for pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Insert pivot table in a new worksheet
Dim pivotWs As Worksheet
Dim pivotTable As PivotTable
Dim pivotCache As PivotCache

Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
Set pivotWs = ThisWorkbook.Worksheets.Add
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="PivotTable1")

' Add row fields to pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add data field to pivot table
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
End With

' Get the data field from the pivot table
Dim dataFieldValue As Variant
dataFieldValue = pivotTable.PivotFields("Sum of Price").DataRange.Cells(1, 1).Value

' Set label and value in pivot worksheet
pivotWs.Range("A12").Value = "Data field value"          ' Set cell A12
pivotWs.Range("B12").Value = dataFieldValue            ' Set cell B12 with data field value

' Update the data field label
pivotTable.PivotFields("Sum of Price").Name = "My Sum of Price" ' Rename data field
pivotWs.Range("A13").Value = "New Data field value"           ' Set cell A13
pivotWs.Range("B13").Value = pivotTable.PivotFields("My Sum of Price").DataRange.Cells(1, 1).Value ' Set cell B13 with updated data field value
```