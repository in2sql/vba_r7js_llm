## Description / Описание

**English:**  
This code initializes an Excel worksheet by setting up headers and populating data in specific cells. It then defines a data range, creates a pivot table on a new worksheet based on that data, adds row fields and a data field to the pivot table, and finally sets and retrieves the layout configuration for the 'Region' field in the pivot table.

**Русский:**  
Этот код инициализирует рабочий лист Excel, устанавливая заголовки и заполняя данные в определенных ячейках. Затем он определяет диапазон данных, создает сводную таблицу на новом листе на основе этих данных, добавляет строковые поля и поле данных в сводную таблицу, и, наконец, устанавливает и получает конфигурацию макета для поля 'Region' в сводной таблице.

```vba
' VBA Code Equivalent to OnlyOffice JS Example

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set header values
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Populate Region data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Populate Style data
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Populate Price data
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the PivotTable
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the PivotTable
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A3"), TableName:="PivotTable1")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set layout for 'Region' field
    pivotSheet.Range("A12").Value = "Region layout compact"
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Region")
    pivotSheet.Range("B12").Value = pivotField.LayoutCompactRow
End Sub
```

```js
// OnlyOffice JS Code Equivalent

function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set header values
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');
    
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
    
    // Define the data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert a new worksheet with PivotTable
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get the pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the 'Region' pivot field
    var pivotField = pivotTable.GetPivotFields('Region');
    
    // Set layout configuration for 'Region'
    pivotWorksheet.GetRange('A12').SetValue('Region layout compact');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutCompactRow());
}
```