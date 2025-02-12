# Description / Описание

**English:** This code populates an Excel worksheet with region, style, and price data, creates a pivot table based on this data, configures the pivot table's layout, and displays the layout configuration status.

**Russian:** Этот код заполняет лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, настраивает макет сводной таблицы и отображает статус конфигурации макета.

## VBA Code / Код VBA

```vba
' Populate worksheet with data, create a pivot table, configure its layout, and display layout status

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region" ' Region
    ws.Range("C1").Value = "Style"  ' Style
    ws.Range("D1").Value = "Price"  ' Price
    
    ' Populate Region
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Populate Style
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Populate Price
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Configure the layout for 'Region' field
    Set pivotField = pivotTable.PivotFields("Region")
    pivotField.LayoutCompactRow = False
    
    ' Display the layout configuration status
    pivotWs.Range("A12").Value = "Region layout compact"
    pivotWs.Range("B12").Value = pivotField.LayoutCompactRow
End Sub
```

## OnlyOffice JS Code / Код OnlyOffice JS

```javascript
// Populate worksheet with data, create a pivot table, configure its layout, and display layout status

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Region
oWorksheet.GetRange('C1').SetValue('Style');  // Style
oWorksheet.GetRange('D1').SetValue('Price');  // Price

// Populate Region
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate Style
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the active sheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set layout compact row to false
pivotField.SetLayoutCompactRow(false);

// Display the layout configuration status
pivotWorksheet.GetRange('A12').SetValue('Region layout compact');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutCompactRow());
```