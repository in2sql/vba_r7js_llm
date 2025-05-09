# Description
This script populates an Excel worksheet with regional sales data, creates a pivot table based on that data, and then lists the unique style items from the pivot table on a new worksheet.

Этот скрипт заполняет рабочий лист Excel данными о продажах по регионам, создает сводную таблицу на основе этих данных и затем выводит уникальные элементы стиля из сводной таблицы на новом рабочем листе.

```vba
' VBA code to populate data, create a pivot table, and list pivot items

Sub CreatePivotTableAndListItems()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim pivotField As PivotField
    Dim pivotItem As PivotItem
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
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
    Set dataRange = ws.Range("$B$1:$D$5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add fields to the pivot table
    With pvt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount
    End With
    
    ' Get the Style pivot field
    Set pivotField = pvt.PivotFields("Style")
    
    ' Add header for style item captions
    pivotWs.Cells(15, 1).Value = "Style item captions"
    
    ' List all unique Style items
    i = 1
    For Each pivotItem In pivotField.PivotItems
        pivotWs.Cells(15 + i, 2).Value = pivotItem.Name
        i = i + 1
    Next pivotItem
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to populate data, create a pivot table, and list pivot items

function createPivotAndListItems() {
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
    
    // Add fields to the pivot table
    pivotTable.AddFields({
        columns: ['Style'],
        rows: 'Region',
    });
    
    // Add data field
    pivotTable.AddDataField('Style');
    
    // Get the pivot table worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the Style pivot field
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Get all pivot items in the Style field
    var pivotItems = pivotField.GetPivotItems();
    
    // Set header for style item captions
    pivotWorksheet.GetRangeByNumber(15, 0).SetValue('Style item captions');
    
    // List each pivot item's caption
    for (var i = 0; i < pivotItems.length; i += 1) {
        pivotWorksheet.GetRangeByNumber(15 + i, 1).SetValue(pivotItems[i].GetCaption());
    }
}
```