**Description / Описание**

This script populates an Excel worksheet with region, style, and price data, then creates a pivot table summarizing the prices by region and style.

Этот скрипт заполняет рабочий лист Excel данными о регионе, стиле и цене, а затем создает сводную таблицу, суммирующую цены по регионам и стилям.

---

```vba
' VBA Code to Populate Data and Create Pivot Table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
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
    
    ' Define data range for Pivot Table
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for Pivot Table
    Dim pvtWs As Worksheet
    Set pvtWs = ThisWorkbook.Worksheets.Add
    pvtWs.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Dim pvtCache As PivotCache
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create Pivot Table
    Dim pvt As PivotTable
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=pvtWs.Range("A3"), TableName:="PivotTable1")
    
    ' Add Row Fields
    With pvt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add Data Field
    pvt.AddDataField pvt.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set layout to show blank lines
    pvt.RowAxisLayout xlTabularRow
End Sub
```

```javascript
// OnlyOffice JS Code to Populate Data and Create Pivot Table

// Get the active worksheet
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

// Define the data range for the Pivot Table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert Pivot Table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Row Fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add Data Field
pivotTable.AddDataField('Price');

// Set layout to show blank lines
pivotTable.SetLayoutBlankLine(true);
```