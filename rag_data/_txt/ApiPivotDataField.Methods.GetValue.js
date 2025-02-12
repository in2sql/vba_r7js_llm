### Description
This script creates a pivot table from a dataset in an Excel worksheet, sets headers, populates data, adds row and data fields to the pivot table, and retrieves a value from the pivot table to display in a specific cell.

Этот скрипт создает сводную таблицу из набора данных на листе Excel, устанавливает заголовки, заполняет данные, добавляет строковые и данные поля в сводную таблицу и извлекает значение из сводной таблицы для отображения в определенной ячейке.

```javascript
// JavaScript (OnlyOffice) Code
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

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add a data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active sheet where the pivot table is located
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field from the pivot table
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in specific cells with data from the pivot table
pivotWorksheet.GetRange('A12').SetValue('The Data field value');
pivotWorksheet.GetRange('B12').SetValue(dataField.GetValue());
```

```vba
' VBA Code
Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"

    ' Populate data
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
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")

    ' Create a new worksheet for the pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = Worksheets.Add

    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)

    ' Add row fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Retrieve value from the pivot table and set it in specific cells
    pivotWs.Range("A12").Value = "The Data field value"
    pivotWs.Range("B12").Value = pivotTable.PivotFields("Sum of Price").DataRange.Cells(1, 1).Value
End Sub
```