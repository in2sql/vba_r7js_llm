---

**Description / Описание**

**English:**  
The code initializes a worksheet with headers and data for Region, Style, and Price. It then creates a pivot table on a new worksheet, adds Style as a page field and Region as a row field, adds Style as a data field, and finally displays the current page of the Style field in cells A13 and B13.

**Russian:**  
Код инициализирует рабочий лист с заголовками и данными для Регион, Стиль и Цена. Затем он создает сводную таблицу на новом листе, добавляет Стиль в качестве полевого фильтра и Регион в качестве строкового поля, добавляет Стиль в качестве поля данных и, наконец, отображает текущую страницу поля Стиля в ячейках A13 и B13.

```vba
' VBA Code to replicate the OnlyOffice JS functionality

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
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data
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
    
    ' Define data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Sheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlPageField
        .PivotFields("Style").Position = 1
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount
    End With
    
    ' Get the Style pivot field
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Set values in pivot worksheet
    pivotWs.Range("A13").Value = "Current Page"
    pivotWs.Range("B13").Value = pivotField.CurrentPage
End Sub
```

```javascript
// OnlyOffice JS Code to create and manipulate pivot table

function createPivot() {
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
    
    // Insert pivot table on a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add fields to pivot table
    pivotTable.AddFields({
        pages: ['Style'], // Add Style as page field
        rows: 'Region'    // Add Region as row field
    });
    
    // Add Style as data field
    pivotTable.AddDataField('Style', 'Count of Style', 'count');
    
    // Get the active sheet for the pivot table
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the Style pivot field
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Set values in pivot worksheet
    pivotWorksheet.GetRange('A13').SetValue('Current Page');
    pivotWorksheet.GetRange('B13').SetValue(pivotField.GetCurrentPage());
}
```