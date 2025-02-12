**Description / Описание**

This code creates a worksheet, sets up headers and data in columns B, C, and D, then creates a pivot table with specified rows and columns, adds a data field, and modifies the pivot table fields accordingly.
Этот код создает лист, устанавливает заголовки и данные в столбцах B, C и D, затем создает сводную таблицу с заданными строками и столбцами, добавляет поле данных и изменяет поля сводной таблицы соответствующим образом.

```vba
' VBA Code Equivalent to OnlyOffice API Example

Sub CreatePivotTable()
    ' Define variables
    Dim oWorksheet As Worksheet
    Dim dataRef As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotWorksheet As Worksheet
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define data range for pivot table
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRef)
    
    ' Add a new worksheet for pivot table
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWorksheet.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlColumnField
        .PivotFields("Region").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Get the 'Region' pivot field
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Disable dragging 'Region' to columns by ensuring it's only in rows
    pivotField.Orientation = xlRowField
    
    ' Set values in the pivot sheet
    pivotWorksheet.Range("A13").Value = "Drag to column"
    pivotWorksheet.Range("B13").Value = pivotField.Orientation <> xlColumnField
    pivotWorksheet.Range("A14").Value = "Try drag Region to columns!"
End Sub
```

```javascript
// OnlyOffice JavaScript API Example

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

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    columns: ['Style'],
    rows: 'Region',
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Disable dragging 'Region' to columns
pivotField.SetDragToColumn(false);

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to column');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToColumn());
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to columns!');
```