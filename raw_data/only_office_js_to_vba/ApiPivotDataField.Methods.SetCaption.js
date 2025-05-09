**Description:**
This code sets values in specific cells of a worksheet, creates a pivot table based on a data range, adds fields to the pivot table, and modifies the data field caption.
Этот код устанавливает значения в определенных ячейках листа, создает сводную таблицу на основе диапазона данных, добавляет поля в сводную таблицу и изменяет заголовок поля данных.

```vba
' VBA Code
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim pivotWs As Worksheet
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataField As PivotField
    
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
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    Set dataField = pivotTable.PivotFields("Price")
    dataField.Orientation = xlDataField
    dataField.Function = xlSum
    dataField.Name = "Sum of Price"
    
    ' Set data field caption
    pivotWs.Range("A12").Value = "Data field caption"
    pivotWs.Range("B12").Value = dataField.Name
    
    dataField.Caption = "My Sum of Price"
    pivotWs.Range("A13").Value = "New Data field caption"
    pivotWs.Range("B13").Value = dataField.Caption
End Sub
```

```javascript
// OnlyOffice JS Code
function createPivotTable() {
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
    
    // Define the data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert pivot table in new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get the pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    var dataField = pivotTable.GetDataFields('Sum of Price');
    
    // Set data field caption
    pivotWorksheet.GetRange('A12').SetValue('Data field caption');
    pivotWorksheet.GetRange('B12').SetValue(dataField.GetCaption());
    
    dataField.SetCaption('My Sum of Price');
    pivotWorksheet.GetRange('A13').SetValue('New Data field caption');
    pivotWorksheet.GetRange('B13').SetValue(dataField.GetCaption());
}
```