# Code Description

**English:** This script populates a worksheet with sample data, creates a pivot table based on that data, adds specific fields to the pivot table, and retrieves the caption of the 'Style' field in the pivot table.

**Russian:** Этот скрипт заполняет рабочий лист примерными данными, создает сводную таблицу на основе этих данных, добавляет определенные поля в сводную таблицу и извлекает подпись поля 'Style' в сводной таблице.

## VBA Code

```vba
' VBA Macro to populate data, create a pivot table, and retrieve field caption

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

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

    ' Define data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"

    ' Create Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)

    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable")

    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Get the caption of 'Style' field
    pivotWorksheet.Range("A12").Value = "The Style field caption"
    pivotWorksheet.Range("B12").Value = pivotTable.PivotFields("Style").Caption

End Sub
```

## OnlyOffice JavaScript Code

```javascript
// JavaScript to populate data, create a pivot table, and retrieve field caption using OnlyOffice API

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Populate headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set Region East
oWorksheet.GetRange('B3').SetValue('West');   // Set Region West
oWorksheet.GetRange('B4').SetValue('East');   // Set Region East
oWorksheet.GetRange('B5').SetValue('West');   // Set Region West

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set Style Fancy
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set Style Fancy
oWorksheet.GetRange('C4').SetValue('Tee');    // Set Style Tee
oWorksheet.GetRange('C5').SetValue('Tee');    // Set Style Tee

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);     // Set Price 42.5
oWorksheet.GetRange('D3').SetValue(35.2);     // Set Price 35.2
oWorksheet.GetRange('D4').SetValue(12.3);     // Set Price 12.3
oWorksheet.GetRange('D5').SetValue(24.8);     // Set Price 24.8

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Add Region and Style as row fields
});

// Add Price as data field
pivotTable.AddDataField('Price');  // Add Price as data field

// Get the active sheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set caption descriptions
pivotWorksheet.GetRange('A12').SetValue('The Style field caption'); // Set label
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetCaption());   // Set caption value
```