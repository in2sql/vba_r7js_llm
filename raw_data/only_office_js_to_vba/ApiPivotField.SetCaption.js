# Create and Manipulate a Pivot Table in Excel / Создание и управление сводной таблицей в Excel

This script sets up data in an Excel worksheet, creates a pivot table based on that data, and modifies pivot table fields.  
Этот скрипт заполняет данными лист Excel, создает сводную таблицу на основе этих данных и изменяет поля сводной таблицы.

```vba
' VBA Code to create and manipulate a pivot table

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
    
    ' Set data values
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
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Set row layout to tabular
    pivotTable.RowAxisLayout xlTabularRow
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Modify pivot table fields
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Set captions
    pivotWs.Range("A12").Value = "Style field caption"
    pivotWs.Range("B12").Value = pivotField.Caption
    
    pivotWs.Range("A14").Value = "New Style field caption"
    pivotField.Caption = "My caption"
    pivotWs.Range("B14").Value = pivotField.Caption
End Sub
```

```javascript
// OnlyOffice JS Code to create and manipulate a pivot table

function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');
    
    // Set data values
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
    
    // Insert a new pivot table in a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Set row axis layout to tabular
    pivotTable.SetRowAxisLayout("Tabular", false);
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get the pivot table worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the 'Style' pivot field
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Set captions
    pivotWorksheet.GetRange('A12').SetValue('Style field caption');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetCaption());
    
    pivotWorksheet.GetRange('A14').SetValue('New Style field caption');
    pivotField.SetCaption('My caption');
    pivotWorksheet.GetRange('B14').SetValue(pivotField.GetCaption());
}
```