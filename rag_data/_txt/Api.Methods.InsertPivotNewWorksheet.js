**Description / Описание**

This code sets up a worksheet with data in specific cells and creates a pivot table based on that data.
Этот код заполняет лист данными в определённых ячейках и создаёт сводную таблицу на основе этих данных.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set cell B1 to 'Region'
oWorksheet.GetRange('C1').SetValue('Price'); // Set cell C1 to 'Price'

// Set data
oWorksheet.GetRange('B2').SetValue('East'); // Set cell B2 to 'East'
oWorksheet.GetRange('B3').SetValue('West'); // Set cell B3 to 'West'
oWorksheet.GetRange('C2').SetValue(42.5); // Set cell C2 to 42.5
oWorksheet.GetRange('C3').SetValue(35.2); // Set cell C3 to 35.2

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");

// Insert a pivot table in a new worksheet based on the data range
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region" ' Set cell B1 to "Region"
oWorksheet.Range("C1").Value = "Price" ' Set cell C1 to "Price"

' Set data
oWorksheet.Range("B2").Value = "East" ' Set cell B2 to "East"
oWorksheet.Range("B3").Value = "West" ' Set cell B3 to "West"
oWorksheet.Range("C2").Value = 42.5 ' Set cell C2 to 42.5
oWorksheet.Range("C3").Value = 35.2 ' Set cell C3 to 35.2

' Define data range
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:C3") ' Define the data range

' Insert a pivot table in a new worksheet based on the data range
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)

Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add ' Create new worksheet for pivot table

Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")
```