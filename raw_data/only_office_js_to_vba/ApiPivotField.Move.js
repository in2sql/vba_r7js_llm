**Script to populate cells and create a pivot table using Excel VBA and OnlyOffice JavaScript API.  
Скрипт для заполнения ячеек и создания сводной таблицы с использованием Excel VBA и JavaScript API OnlyOffice.**

```vba
' VBA code to populate cells and create a pivot table

Sub CreatePivotTable()
    Dim oSheet As Worksheet
    Dim pRange As Range
    Dim pTable As PivotTable
    Dim pCache As PivotCache
    Dim pField As PivotField

    ' Set the active sheet
    Set oSheet = ActiveSheet

    ' Set headers
    oSheet.Range("B1").Value = "Region"
    oSheet.Range("C1").Value = "Style"
    oSheet.Range("D1").Value = "Price"

    ' Populate Region data
    oSheet.Range("B2").Value = "East"
    oSheet.Range("B3").Value = "West"
    oSheet.Range("B4").Value = "East"
    oSheet.Range("B5").Value = "West"

    ' Populate Style data
    oSheet.Range("C2").Value = "Fancy"
    oSheet.Range("C3").Value = "Fancy"
    oSheet.Range("C4").Value = "Tee"
    oSheet.Range("C5").Value = "Tee"

    ' Populate Price data
    oSheet.Range("D2").Value = 42.5
    oSheet.Range("D3").Value = 35.2
    oSheet.Range("D4").Value = 12.3
    oSheet.Range("D5").Value = 24.8

    ' Define the data range
    Set pRange = oSheet.Range("B1:D5")

    ' Create Pivot Cache
    Set pCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pRange)

    ' Create Pivot Table on a new worksheet
    Set pTable = pCache.CreatePivotTable(TableDestination:=ActiveWorkbook.Worksheets.Add.Range("A3"), TableName:="PivotTable1")

    ' Add row and column fields to Pivot Table
    pTable.PivotFields("Region").Orientation = xlRowField
    pTable.PivotFields("Style").Orientation = xlColumnField

    ' Add data field to Pivot Table
    pTable.AddDataField pTable.PivotFields("Price"), "Sum of Price", xlSum

    ' Set a message in cell A10
    oSheet.Range("A10").Value = "The Region field will be moved soon"

    ' Schedule the movement of the 'Region' field after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "MoveRegionField"
End Sub

' Subroutine to move Pivot Field
Sub MoveRegionField()
    Dim pTable As PivotTable
    Set pTable = ActiveSheet.PivotTables(1)
    pTable.PivotFields("Region").Orientation = xlColumnField
End Sub
```

```javascript
// JavaScript code to populate cells and create a pivot table using OnlyOffice API

var oWorksheet = Api.GetActiveSheet();

// Set headers
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert Pivot Table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to Pivot Table
pivotTable.AddFields({
    rows: 'Region',
    columns: 'Style',
});

// Get active sheet for pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Add data field to Pivot Table
pivotTable.AddDataField('Price');

// Get pivot field 'Region'
var pivotField = pivotTable.GetPivotFields('Region');

// Set a message in cell A10
pivotWorksheet.GetRange('A10').SetValue('The Region field will be moved soon');

// Move 'Region' field to Columns after 5 seconds
setTimeout(function () {
    pivotField.Move('Columns');
}, 5000);
```