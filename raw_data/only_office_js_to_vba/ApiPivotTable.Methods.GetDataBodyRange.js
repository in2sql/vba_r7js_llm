**Description / Описание**

This script populates an Excel worksheet with region, style, and price data, then creates a pivot table based on this data.  
Этот скрипт заполняет рабочий лист Excel данными о регионе, стиле и цене, затем создает сводную таблицу на основе этих данных.

```vba
' VBA Code to populate worksheet and create a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
    ws.Range("B1").Value = "Region" ' Set header for Region
    ws.Range("C1").Value = "Style"  ' Set header for Style
    ws.Range("D1").Value = "Price"  ' Set header for Price
    
    ' Populate Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Populate Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Populate Price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Create Pivot Cache
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Add a new worksheet for the pivot table
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "PivotTableSheet"
    
    ' Create Pivot Table
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=ws.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Set pivot table fields
    With pt
        .PivotFields("Region").Orientation = xlRowField ' Add Region to Rows
        .PivotFields("Style").Orientation = xlColumnField ' Add Style to Columns
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Add Price to Values
    End With
    
    ' Select the data body range of the pivot table
    pt.DataBodyRange.Select
End Sub
```

```javascript
// JavaScript Code to populate worksheet and create a pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

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

// Insert a new worksheet with the pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
	rows: 'Region',   // Add Region to Rows
	columns: 'Style'  // Add Style to Columns
});

// Add Price as a data field
pivotTable.AddDataField('Price', 'Sum of Price', 'Sum');

// Select the data body range of the pivot table
pivotTable.GetDataBodyRange().Select();
```