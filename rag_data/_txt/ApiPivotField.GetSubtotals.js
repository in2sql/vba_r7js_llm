# Description / Описание

**English:**  
This script populates an Excel worksheet with data, creates a pivot table based on that data, and then lists the subtotals for each region.

**Русский:**  
Этот скрипт заполняет рабочий лист Excel данными, создает сводную таблицу на основе этих данных, а затем выводит промежуточные итоги по каждому региону.

```vba
' VBA Code to populate data, create a pivot table, and list region subtotals

Sub CreatePivotTableAndListSubtotals()
    Dim oWorksheet As Worksheet
    Dim pivotWorksheet As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    Dim subtotals As Variant
    Dim i As Integer
    Dim k As Integer

    ' Set the active worksheet
    Set oWorksheet = ActiveSheet

    ' Populate headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"

    ' Populate Region data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"

    ' Populate Style data
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"

    ' Populate Price data
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8

    ' Define the data range
    Set dataRange = oWorksheet.Range("B1:D5")

    ' Create a pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Add a new worksheet for the pivot table
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"

    ' Create the pivot table
    Set pivotTable = pivotWorksheet.PivotTables.Add( _
        PivotCache:=pivotCache, _
        TableDestination:=pivotWorksheet.Range("A1"), _
        TableName:="SalesPivotTable")

    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlColumnField
        .PivotFields("Region").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Get the pivot field for Region
    Set pivotField = pivotTable.PivotFields("Region")

    ' Get subtotals for each Region
    subtotals = pivotField.DataRange.Cells
    pivotWorksheet.Range("A11").Value = "Region subtotals"
    k = 12
    For i = 1 To pivotField.PivotItems.Count
        pivotWorksheet.Cells(k, 1).Value = pivotField.PivotItems(i).Name
        pivotWorksheet.Cells(k, 2).Value = pivotField.PivotItems(i).DataRange.Cells(1, 1).Value
        k = k + 1
    Next i
End Sub
```

```js
// JavaScript Code to populate data, create a pivot table, and list region subtotals using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers in B1, C1, D1
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Style' as column field and 'Region' as row field
pivotTable.AddFields({
	columns: ['Style'],
	rows: 'Region',
});

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field for 'Region'
var pivotField = pivotTable.GetPivotFields('Region');

// Get subtotals from the 'Region' pivot field
var subtotals = pivotField.GetSubtotals();

// Set the header for subtotals
pivotWorksheet.GetRange('A11').SetValue('Region subtotals');

// Initialize row counter
let k = 12;

// Iterate through subtotals and set values in the worksheet
for (var i in subtotals) {
	pivotWorksheet.GetRangeByNumber(k, 0).SetValue(i);
	pivotWorksheet.GetRangeByNumber(k, 1).SetValue(subtotals[i]);
	k++;
}
```