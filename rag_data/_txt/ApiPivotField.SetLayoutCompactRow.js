# Conversion of OnlyOffice API JavaScript Code to Excel VBA
# Конвертация кода OnlyOffice API JavaScript в Excel VBA

**Description:**  
This code initializes a worksheet, sets up data, creates a pivot table, modifies its layout, and displays layout settings.

**Описание:**  
Этот код инициализирует лист, устанавливает данные, создает сводную таблицу, изменяет ее макет и отображает настройки макета.

```javascript
// JavaScript (OnlyOffice) Code

// Get the active worksheet
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

// Insert pivot table in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Modify pivot field layout
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.SetLayoutCompactRow(false);

// Display layout settings
pivotWorksheet.GetRange('A12').SetValue('Region layout compact');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutCompactRow());
```

```vba
' VBA Code

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

' Add pivot table to new worksheet
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add
pivotSheet.Name = "PivotTableSheet"

Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotSheet.Range("A3"), _
    TableName:="PivotTable1")

' Add fields to pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    .PivotFields("Price").Orientation = xlDataField
    .PivotFields("Price").Function = xlSum
End With

' Modify pivot field layout
With pivotTable.PivotFields("Region")
    .LayoutCompact = False
End With

' Display layout settings
pivotSheet.Range("A12").Value = "Region layout compact"
pivotSheet.Range("B12").Value = pivotTable.PivotFields("Region").LayoutCompact
```