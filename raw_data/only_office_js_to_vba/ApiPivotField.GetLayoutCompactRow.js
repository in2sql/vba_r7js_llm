# Description / Описание

This code sets values in specified cells, creates a pivot table from a data range, adds row and data fields, and sets layout options.
Этот код устанавливает значения в указанные ячейки, создает сводную таблицу из диапазона данных, добавляет строки и поля данных, и настраивает параметры макета.

## JavaScript (OnlyOffice API) Code

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
// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet containing the pivot table
var pivotWorksheet = Api.GetActiveSheet();
// Get the 'Region' field from the pivot table
var pivotField = pivotTable.GetPivotFields('Region');

// Set layout option for the 'Region' field
pivotWorksheet.GetRange('A12').SetValue('Region layout compact');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutCompactRow());
```

## Excel VBA Code

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Set Region values
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Set Style values
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Set Price values
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotWorksheet As Worksheet
Set pivotWorksheet = ThisWorkbook.Worksheets.Add
pivotWorksheet.Name = "PivotTableSheet"

' Create the pivot cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRef)

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotWorksheet.Range("A1"), _
    TableName:="PivotTable1")

' Add row fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add data field to the pivot table
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
End With

' Set layout option for the 'Region' field
pivotWorksheet.Range("A12").Value = "Region layout compact"
pivotTable.RowAxisLayout xlCompactRow
```