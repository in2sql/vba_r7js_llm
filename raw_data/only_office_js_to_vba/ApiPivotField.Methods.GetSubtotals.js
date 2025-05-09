### Description / Описание
**English:** This code populates an Excel worksheet with region, style, and price data, creates a pivot table based on this data, and then lists the subtotals for each region.

**Русский:** Этот код заполняет рабочий лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, а затем выводит подытоги по каждому региону.

```vba
' VBA Code to populate worksheet, create pivot table, and list subtotals

Sub CreatePivotTableAndSubtotals()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    Dim subtotals As Variant
    Dim i As Integer

    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"

    ' Populate data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"

    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"

    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8

    ' Define data range
    Set dataRange = ws.Range("B1:D5")

    ' Create a new worksheet for pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"

    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")

    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlColumnField
        .PivotFields("Region").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Get subtotals for 'Region' field
    Set pivotField = pivotTable.PivotFields("Region")
    ' Assuming subtotals are the sum for each region
    ' This part may need customization based on actual subtotal requirements

    ' List subtotals below the pivot table
    pivotWs.Range("A11").Value = "Region subtotals"
    i = 12
    Dim cell As Range
    For Each cell In pivotField.DataRange
        pivotWs.Cells(i, 1).Value = cell.Value
        pivotWs.Cells(i, 2).Value = cell.Offset(0, 1).Value ' Assuming subtotal is next column
        i = i + 1
    Next cell
End Sub
```

```javascript
// JavaScript Code to populate worksheet, create pivot table, and list subtotals using OnlyOffice API

var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    columns: ['Style'],
    rows: 'Region',
});
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Get subtotals for 'Region'
var subtotals = pivotField.GetSubtotals();

// List subtotals below the pivot table
pivotWorksheet.GetRange('A11').SetValue('Region subtotals');
let k = 12;
for (var i in subtotals) {
    pivotWorksheet.GetRangeByNumber(k, 0).SetValue(i);
    pivotWorksheet.GetRangeByNumber(k++, 1).SetValue(subtotals[i]);
}
```