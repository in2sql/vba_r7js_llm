### Description

**English:** The script populates specific cells with data, creates a pivot table from the data, adds row and data fields to the pivot table, inserts a message in a cell, and changes the orientation of a pivot field after a delay.

**Russian:** Скрипт заполняет определенные ячейки данными, создает сводную таблицу из этих данных, добавляет поля строк и данных в сводную таблицу, вставляет сообщение в ячейку и изменяет ориентацию поля сводной таблицы через задержку.

### OnlyOffice JavaScript Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
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

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add Price as a data field
pivotTable.AddDataField('Price');

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Insert a message in cell A12
pivotWorksheet.GetRange('A12').SetValue('The Style field orientation will change soon');

// Get the Style field from the pivot table
var pivotField = pivotTable.GetPivotFields('Style');

// After 5 seconds, change the orientation of the Style field to Columns
setTimeout(function () {
    pivotField.SetOrientation("Columns");
}, 5000);
```

### Excel VBA Code

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
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

' Insert a new pivot table on a new worksheet
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add
Dim pivotTable As PivotTable
Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)

' Add row fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    ' Add Price as a data field
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
End With

' Insert a message in cell A12
pivotSheet.Range("A12").Value = "The Style field orientation will change soon"

' Change the orientation of the Style field to Columns after 5 seconds
Application.OnTime Now + TimeValue("00:00:05"), "ChangeStyleOrientation"

' Subroutine to change the orientation of the Style field
Sub ChangeStyleOrientation()
    With pivotTable.PivotFields("Style")
        .Orientation = xlColumnField
    End With
End Sub
```