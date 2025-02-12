**Description:**

English: This code populates an Excel worksheet with data, creates a pivot table on a new worksheet, adds 'Region' and 'Style' as row fields, 'Price' as a data field, adjusts the subtotal location for 'Region' to the bottom, renames the subtotal, and sets specific cell values based on the pivot table's subtotal.

Russian: Этот код заполняет рабочий лист Excel данными, создает сводную таблицу на новом листе, добавляет 'Region' и 'Style' как поля строк, 'Price' как поле данных, корректирует расположение подитогов для 'Region' внизу, переименовывает подитог и устанавливает значения определенных ячеек на основе подитога сводной таблицы.

```vba
' VBA code to populate worksheet, create pivot table, and configure fields

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long
    
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
    
    ' Define data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for pivot table
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
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Style").Position = 2
    End With
    
    ' Add data field
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Price"
    End With
    
    ' Adjust subtotal location for 'Region' to bottom
    With pivotTable.PivotFields("Region")
        .Subtotals(1) = False ' Disable default subtotals
        .Subtotals(1) = True ' Enable custom subtotal
        .LayoutSubtotal(1) = xlAtBottom
        .Name = "Region Subtotal Name"
    End With
    
    ' Set specific cell values based on pivot table's subtotal
    pivotWs.Range("A14").Value = "Region subtotal name"
    pivotWs.Range("B14").Value = pivotTable.PivotFields("Region").Name
End Sub
```

```javascript
// OnlyOffice JS code to populate worksheet, create pivot table, and configure fields

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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Configure subtotal for 'Region'
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.SetLayoutSubtotalLocation('Bottom');

pivotField.SetSubtotalName('My name');

// Set specific cell values based on pivot table's subtotal
pivotWorksheet.GetRange('A14').SetValue('Region subtotal name');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetSubtotalName());
```