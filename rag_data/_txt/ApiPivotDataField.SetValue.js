**Description / Описание:**  
This script initializes a worksheet, sets specific cell values, creates a pivot table, and modifies data fields.  
Этот скрипт инициализирует рабочий лист, устанавливает значения определенных ячеек, создает сводную таблицу и изменяет поля данных.

```vba
' VBA Code to mimic OnlyOffice API functionality
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
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
    
    ' Define data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    Set dataField = pivotTable.PivotFields("Price")
    dataField.Orientation = xlDataField
    dataField.Function = xlSum
    dataField.Name = "Sum of Price"
    
    ' Set values in pivot worksheet
    pivotWs.Range("A12").Value = "Data field value"
    pivotWs.Range("B12").Value = dataField.Value
    
    ' Modify data field name
    dataField.Name = "My Sum of Price"
    pivotWs.Range("A13").Value = "New Data field value"
    pivotWs.Range("B13").Value = dataField.Value
End Sub
```

```javascript
// OnlyOffice JS Code to create and manipulate a pivot table
function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();

    // Set header values
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

    // Add row fields to the pivot table
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });

    // Add data field to the pivot table
    pivotTable.AddDataField('Price');

    // Get the active worksheet where pivot table is placed
    var pivotWorksheet = Api.GetActiveSheet();

    // Get the data field value from the pivot table
    var dataField = pivotTable.GetDataFields('Sum of Price');

    // Set values in the pivot worksheet
    pivotWorksheet.GetRange('A12').SetValue('Data field value');
    pivotWorksheet.GetRange('B12').SetValue(dataField.GetValue());

    // Modify the data field name
    dataField.SetValue('My Sum of Price');
    pivotWorksheet.GetRange('A13').SetValue('New Data field value');
    pivotWorksheet.GetRange('B13').SetValue(dataField.GetValue());
}
```