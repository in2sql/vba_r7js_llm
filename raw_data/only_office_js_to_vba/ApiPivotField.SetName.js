**Description / Описание**

This code sets up a worksheet with specified values, creates a pivot table, adds fields and data, modifies field names, and outputs them to specific cells.

Этот код заполняет рабочий лист заданными значениями, создаёт сводную таблицу, добавляет поля и данные, изменяет названия полей и выводит их в определённые ячейки.

```vba
' VBA Code Equivalent
Sub CreatePivotTable()
    ' Set reference to the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set Region values
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set Style values
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set Price values
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .RowAxisLayout xlTabularRow
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Rename a pivot field
    pivotTable.PivotFields("Style").Name = "My name"
    
    ' Output field names to specific cells
    pivotWs.Range("A12").Value = "Style field name"
    pivotWs.Range("B12").Value = pivotTable.PivotFields("Style").Name
    
    pivotWs.Range("A14").Value = "New Style field name"
    pivotWs.Range("B14").Value = pivotTable.PivotFields("My name").Name
End Sub
```

```javascript
// OnlyOffice JS Code
// This script sets up a worksheet with data, creates a pivot table, adds fields, modifies them, and outputs their names to specific cells.

var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set Region values
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set Style values
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price values
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range and insert pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});
pivotTable.SetRowAxisLayout("Tabular", false);

// Add data field
pivotTable.AddDataField('Price');

// Get pivot worksheet and modify fields
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Style');

// Output initial field name
pivotWorksheet.GetRange('A12').SetValue('Style field name');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetName());

// Rename the field and output new name
pivotWorksheet.GetRange('A14').SetValue('New Style field name');
pivotField.SetName('My name');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetName());
```