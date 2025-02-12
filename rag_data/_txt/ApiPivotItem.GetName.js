**Description:**

*English:*  
This script initializes the active worksheet, sets values in specific cells, creates a new pivot table from a data range, adds specific fields to the pivot table, and populates pivot item names into the worksheet.

*Russian:*  
Этот скрипт инициализирует активный лист, устанавливает значения в определенные ячейки, создает новую сводную таблицу из диапазона данных, добавляет определенные поля в сводную таблицу и заполняет имена элементов сводной таблицы в лист.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set region values
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set style values
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set price values
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    columns: ['Style'],
    rows: 'Region',
});

// Add data field
pivotTable.AddDataField('Style');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get pivot field and items
var pivotField = pivotTable.GetPivotFields('Style');
var pivotItems = pivotField.GetPivotItems();

// Set header for pivot items
pivotWorksheet.GetRangeByNumber(15, 0).SetValue('Style item names');

// Populate pivot item names
for (var i = 0; i < pivotItems.length; i += 1) {
    pivotWorksheet.GetRangeByNumber(15 + i, 1).SetValue(pivotItems[i].GetName());
}
```

```vba
' VBA code equivalent

' Description: Initializes the active worksheet, sets values in cells, creates a pivot table, 
' adds fields and data, and populates pivot item names.

Sub CreatePivotTable()
    Dim oWorksheet As Worksheet
    Dim pivotWorksheet As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    Dim pivotItem As PivotItem
    Dim i As Integer
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set region values
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Set style values
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Set price values
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Create a pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Add a new worksheet for the pivot table
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWorksheet.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount
    End With
    
    ' Populate pivot item names
    Set pivotField = pivotTable.PivotFields("Style")
    pivotWorksheet.Cells(15, 1).Value = "Style item names"
    
    i = 0
    For Each pivotItem In pivotField.PivotItems
        pivotWorksheet.Cells(15 + i, 2).Value = pivotItem.Name
        i = i + 1
    Next pivotItem
End Sub
```