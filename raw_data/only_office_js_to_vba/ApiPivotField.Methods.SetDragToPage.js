# Code Description / Описание кода

This code initializes data in an Excel worksheet, creates a pivot table based on the data, and modifies pivot table settings.

Этот код инициализирует данные на листе Excel, создает сводную таблицу на основе данных и изменяет настройки сводной таблицы.

```vba
' VBA code equivalent to the provided OnlyOffice JS code

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create Pivot Table
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add fields
    With pivotTable
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Region").Orientation = xlColumnField
        .PivotFields("Price").Orientation = xlDataField
    End With
    
    ' Modify pivot field
    pivotTable.PivotFields("Region").EnableItemSelection = False
    
    ' Set values in the pivot worksheet
    pivotWorksheet.Range("A13").Value = "Drag to page"
    pivotWorksheet.Range("B13").Value = pivotTable.PivotFields("Region").EnableItemSelection
    pivotWorksheet.Range("A14").Value = "Try drag Region to pages!"
End Sub
```

```javascript
// OnlyOffice JS code: Initializes data, creates a pivot table, and modifies pivot table settings.

var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data
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

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
	rows: ['Style'],
	columns: 'Region',
});

// Add data field
pivotTable.AddDataField('Price');

// Get pivot worksheet and field
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Region');

// Modify pivot field settings
pivotField.SetDragToPage(false);

// Set values in pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to page');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToPage());
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to pages!');
```