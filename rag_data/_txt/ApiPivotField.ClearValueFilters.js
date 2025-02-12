# Create and manipulate worksheet, insert pivot table / Создание и манипуляция листом, вставка сводной таблицы

```vba
' VBA code to set up data and create a pivot table

Sub CreateDataAndPivot()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"

    ' Set data for Region
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"

    ' Set data for Style
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"

    ' Set data for Price
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8

    ' Define data range
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")

    ' Create Pivot Cache
    Dim pvtCache As PivotCache
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    ' Add a new worksheet for Pivot Table
    Dim pvtSheet As Worksheet
    Set pvtSheet = ThisWorkbook.Worksheets.Add
    pvtSheet.Name = "PivotTableSheet"

    ' Create Pivot Table
    Dim pvtTable As PivotTable
    Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=pvtSheet.Range("A3"), TableName:="PivotTable1")

    ' Add fields to Pivot Table
    With pvtTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Clear value filters on 'Region' field
    pvtTable.PivotFields("Region").ClearAllFilters
End Sub
```

```javascript
// JavaScript code to set up data and create a pivot table using OnlyOffice API

// Get the active worksheet
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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new worksheet with pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to pivot table
pivotTable.AddFields({
    rows: 'Region',
    columns: 'Style',
});

// Add data field to pivot table
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' field in pivot table and clear value filters
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.ClearValueFilters(); 
```