# Code Description / Описание кода

This code sets up an Excel worksheet by entering data into specific cells, inserting a pivot table, adding fields to it, and setting a value based on the pivot table's properties.

Этот код настраивает лист Excel путем ввода данных в определенные ячейки, вставки сводной таблицы, добавления полей к ней и установки значения на основе свойств сводной таблицы.

```vba
' VBA code to replicate OnlyOffice API functionality

Sub SetupWorksheetAndPivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField

    ' Get the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data for Region
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set data for Style
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set data for Price
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Set pivotTable = pivotWs.PivotTableWizard(TableDestination:=pivotWs.Range("A3"), SourceData:=dataRange)
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the pivot field for Region
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set values related to pivot table layout
    pivotWs.Range("A15").Value = "Page break:"
    pivotWs.Range("B15").Value = pivotField.LayoutPageBreak
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to set up worksheet and pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data for Region
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set data for Style
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set data for Price
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field for Region
var pivotField = pivotTable.GetPivotFields('Region');

// Set values related to pivot table layout
pivotWorksheet.GetRange('A15').SetValue('Page break:');
pivotWorksheet.GetRange('B15').SetValue(pivotField.GetLayoutPageBreak());
```