# Code Description / Описание кода

**English:**  
This code populates an active worksheet with region, style, and price data, creates a pivot table based on this data, and then adds a field to the pivot table for dragging to the page area.

**Russian:**  
Этот код заполняет активный лист данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, а затем добавляет поле в сводную таблицу для перетаскивания в область страницы.

## Excel VBA Code

```vba
' VBA code to populate data and create a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
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
    
    ' Define data range for pivot table
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWS As Worksheet
    Set pivotWS = ThisWorkbook.Worksheets.Add
    pivotWS.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWS.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlRowField ' Add Style as row
        .PivotFields("Region").Orientation = xlColumnField ' Add Region as column
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Add Price as data field
    End With
    
    ' Add "Drag to page" label
    pivotWS.Range("A13").Value = "Drag to page"
    
    ' Note: VBA does not have a direct equivalent for GetDragToPage
    ' Additional implementation may be required based on specific needs
End Sub
```

## OnlyOffice JS Code

```javascript
// JavaScript code to populate data and create a pivot table using OnlyOffice API

function createPivotTable(Api) {
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
    
    // Define data range for pivot table
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert a new pivot table on a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add fields to pivot table
    pivotTable.AddFields({
        rows: ['Style'],      // Add Style as row
        columns: 'Region',    // Add Region as column
    });
    
    pivotTable.AddDataField('Price'); // Add Price as data field
    
    // Get the pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the Region pivot field
    var pivotField = pivotTable.GetPivotFields('Region');
    
    // Add "Drag to page" label and value
    pivotWorksheet.GetRange('A13').SetValue('Drag to page');
    pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToPage());
}
```