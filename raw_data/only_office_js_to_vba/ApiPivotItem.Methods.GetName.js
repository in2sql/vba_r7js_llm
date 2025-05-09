# Script Description
**English:**  
The script creates a worksheet, populates data, creates a pivot table, and outputs the names of the "Style" field items.

**Russian:**  
Скрипт создает рабочий лист, заполняет данные, создает сводную таблицу и выводит имена элементов поля "Style".

```vba
' VBA Code to replicate OnlyOffice JS functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim ptSheet As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    Dim pivotItem As PivotItem
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set Price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for PivotTable
    Set ptSheet = ThisWorkbook.Worksheets.Add
    ptSheet.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=ptSheet.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add fields to Pivot Table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount
    End With
    
    ' Get PivotField "Style"
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Write header for Style items
    ptSheet.Cells(15, 1).Value = "Style item names"
    
    ' Loop through PivotItems and write names
    i = 0
    For Each pivotItem In pivotField.PivotItems
        ptSheet.Cells(15 + i + 1, 2).Value = pivotItem.Name
        i = i + 1
    Next pivotItem
End Sub
```

```javascript
// OnlyOffice JS Code to create worksheet, populate data, create pivot table and list pivot items

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set Region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set Style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert PivotTable in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to PivotTable
pivotTable.AddFields({
    columns: ['Style'],
    rows: 'Region',
});

// Add data field
pivotTable.AddDataField('Style');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get PivotField "Style"
var pivotField = pivotTable.GetPivotFields('Style');

// Get PivotItems
var pivotItems = pivotField.GetPivotItems();

// Set header for Style item names
pivotWorksheet.GetRangeByNumber(15, 0).SetValue('Style item names');

// Loop through PivotItems and write names
for (var i = 0; i < pivotItems.length; i += 1) {
    pivotWorksheet.GetRangeByNumber(15 + i, 1).SetValue(pivotItems[i].GetName());
}
```