## Description:
This code creates and populates a spreadsheet with regional data, then creates a pivot table summarizing the data, and finally retrieves and displays the orientation of the 'Style' field in the pivot table.

Описание:
Этот код создает и заполняет электронную таблицу данными о регионах, затем создает сводную таблицу для суммирования данных и, наконец, получает и отображает ориентацию поля 'Style' в сводной таблице.

```javascript
// Get the active sheet
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    rows: 'Region',
    columns: 'Style',
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot table's active sheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set value in cell A12
pivotWorksheet.GetRange('A12').SetValue('The Style field orientation');

// Set value in cell B12 with the pivot field's orientation
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetOrientation());
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

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

' Define the data range
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Insert a new pivot table on a new worksheet
Dim pivotWorksheet As Worksheet
Set pivotWorksheet = Worksheets.Add

Dim pivotTable As PivotTable
Set pivotTable = pivotWorksheet.PivotTableWizard(TableDestination:=pivotWorksheet.Range("A3"), TableName:="PivotTable1", SourceData:=dataRef)

' Add fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlColumnField
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
End With

' Get the 'Style' pivot field
Dim pivotField As PivotField
Set pivotField = pivotTable.PivotFields("Style")

' Set value in cell A12
pivotWorksheet.Range("A12").Value = "The Style field orientation"

' Set value in cell B12 with the pivot field's orientation
pivotWorksheet.Range("B12").Value = pivotField.Orientation
```