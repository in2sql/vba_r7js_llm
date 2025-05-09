### Description / Описание

**English:** This code initializes an active worksheet, sets up headers and data in specific cells, creates a pivot table based on the data, configures the pivot table fields, and sets values in certain cells regarding the pivot table's row dragging functionality.

**Russian:** Этот код инициализирует активный лист, устанавливает заголовки и данные в определенные ячейки, создает сводную таблицу на основе данных, настраивает поля сводной таблицы и устанавливает значения в определенных ячейках, касающиеся функциональности перетаскивания строк в сводной таблице.

```vba
' VBA Code
Sub CreatePivotTable()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet ' Get active sheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set Region data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Set Style data
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Set Price data
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    
    ' Create pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Region").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Access pivot field 'Region'
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set drag to row as false (Not directly possible in VBA, may require workaround)
    ' VBA does not have a direct property to set drag behavior
    
    ' Set values in pivot worksheet
    pivotSheet.Range("A13").Value = "Drag to row"
    pivotSheet.Range("B13").Value = False ' Assuming drag to row is false
    pivotSheet.Range("A14").Value = "Try drag Region to rows!"
End Sub
```

```javascript
// OnlyOffice JS Code
function createPivotTable() {
    // Get active sheet
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
    
    // Insert pivot table
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add fields to pivot table
    pivotTable.AddFields({
        rows: ['Style'],
        columns: 'Region',
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Access pivot field 'Region'
    var pivotField = pivotTable.GetPivotFields('Region');
    
    // Set drag to row as false
    pivotField.SetDragToRow(false);
    
    // Set values in pivot worksheet
    pivotWorksheet.GetRange('A13').SetValue('Drag to row');
    pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToRow());
    pivotWorksheet.GetRange('A14').SetValue('Try drag Region to rows!');
}
```