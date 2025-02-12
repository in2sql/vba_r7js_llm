**Description:**
This code sets up data in a worksheet, creates a pivot table, and formats it.
Этот код заполняет данные на листе, создает сводную таблицу и форматирует ее.

---

### JavaScript Code

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

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet where the pivot table is located
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field from the pivot table and set its number format
var dataField = pivotTable.GetDataFields('Sum of Price');
dataField.SetNumberFormat('0.00E+00');
```

---

### Excel VBA Code

```vba
' Get the active worksheet
Dim ws As Worksheet
Set ws = ThisWorkbook.ActiveSheet

' Set headers
ws.Range("B1").Value = "Region"
ws.Range("C1").Value = "Style"
ws.Range("D1").Value = "Price"

' Set Region values
ws.Range("B2").Value = "East"
ws.Range("B3").Value = "West"
ws.Range("B4").Value = "East"
ws.Range("B5").Value = "West"

' Set Style values
ws.Range("C2").Value = "Fancy"
ws.Range("C3").Value = "Fancy"
ws.Range("C4").Value = "Tee"
ws.Range("C5").Value = "Tee"

' Set Price values
ws.Range("D2").Value = 42.5
ws.Range("D3").Value = 35.2
ws.Range("D4").Value = 12.3
ws.Range("D5").Value = 24.8

' Define the data range for the pivot table
Dim dataRange As Range
Set dataRange = ws.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotWS As Worksheet
Set pivotWS = ThisWorkbook.Worksheets.Add
pivotWS.Name = "PivotTableSheet"

' Create the pivot cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotWS.Range("A3"), _
    TableName:="PivotTable1")

' Add row fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add data field to the pivot table
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
End With

' Set number format for the data field
pivotTable.DataFields(1).NumberFormat = "0.00E+00"
```