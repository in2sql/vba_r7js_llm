### Description / Описание
**English:** This script populates an Excel worksheet with region, style, and price data, creates a pivot table based on that data, and then displays whether the 'Region' and 'Style' fields are shown in the pivot table's axis.

**Русский:** Этот скрипт заполняет рабочий лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, а затем отображает, показываются ли поля «Region» и «Style» на оси сводной таблицы.

```vba
' VBA Code to replicate the OnlyOffice JS functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim lastRow As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Populate headers
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
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivot")
    
    ' Add 'Region' to Rows
    With pivotTable.PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ' Add 'Price' to Values
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Price"
    End With
    
    ' Display whether 'Region' is shown in axis
    pivotWs.Range("A12").Value = "Region showing in axis"
    pivotWs.Range("B12").Value = pivotTable.PivotFields("Region").Orientation = xlRowField
    
    ' Display whether 'Style' is shown in axis
    pivotWs.Range("A13").Value = "Style showing in axis"
    pivotWs.Range("B13").Value = pivotTable.PivotFields("Style").Orientation = xlRowField
End Sub
```

```javascript
// JavaScript Code to replicate the OnlyOffice functionality

// Get the active worksheet
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Region' as row field
pivotTable.AddFields({
	rows: ['Region'],
});

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Display whether 'Region' is shown in axis
pivotWorksheet.GetRange('A12').SetValue('Region showing in axis');
pivotWorksheet.GetRange('B12').SetValue(pivotTable.GetPivotFields('Region').GetShowingInAxis());

// Display whether 'Style' is shown in axis
pivotWorksheet.GetRange('A13').SetValue('Style showing in axis');
pivotWorksheet.GetRange('B13').SetValue(pivotTable.GetPivotFields('Style').GetShowingInAxis());
```