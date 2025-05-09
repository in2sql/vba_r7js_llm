**Description:**

*English:* This code populates an Excel worksheet with regional sales data, creates a pivot table based on this data, adds specific row and data fields to the pivot table, and retrieves the position of a particular data field within the pivot table.

*Russian:* Этот код заполняет лист Excel данными о продажах по регионам, создает сводную таблицу на основе этих данных, добавляет определенные строковые и числовые поля в сводную таблицу и получает позицию конкретного числового поля в сводной таблице.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers in cells B1, C1, D1
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate column B with regions
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate column C with styles
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate column D with prices
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Region' and 'Style' as row fields in the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add 'Price' as a data field in the pivot table
pivotTable.AddDataField('Price');
var dataField = pivotTable.AddDataField('Price');

// Get the worksheet containing the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Set descriptive text in cells A15 and B15
pivotWorksheet.GetRange('A15').SetValue('Sum of Price2 position:');
pivotWorksheet.GetRange('B15').SetValue(dataField.GetPosition());
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers in cells B1, C1, D1
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Populate column B with regions
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Populate column C with styles
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Populate column D with prices
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("$B$1:$D$5")

' Create a pivot cache based on the data range
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add

' Create the pivot table in the new worksheet
Dim pivotTable As PivotTable
Set pivotTable = pivotSheet.PivotTables.Add(PivotCache:=pivotCache, TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")

' Add 'Region' and 'Style' as row fields in the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add 'Price' as a data field in the pivot table
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    .AddDataField .PivotFields("Price"), "Sum of Price2", xlSum
End With

' Retrieve the position of the second 'Price' data field
Dim dataFieldPosition As Integer
dataFieldPosition = pivotTable.DataFields("Sum of Price2").Position

' Set descriptive text in cells A15 and B15 of the pivot sheet
pivotSheet.Range("A15").Value = "Sum of Price2 position:"
pivotSheet.Range("B15").Value = dataFieldPosition
```