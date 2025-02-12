# Description / Описание

**English:**  
This code manipulates an active worksheet by setting values in specific cells, creates a pivot table based on a defined data range, and configures the pivot table fields accordingly.

**Русский:**  
Этот код манипулирует активным листом, устанавливая значения в определенные ячейки, создает сводную таблицу на основе заданного диапазона данных и настраивает поля сводной таблицы соответственно.

## VBA Code

```vba
' VBA code to manipulate worksheet and create a pivot table

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set Region values
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Set Style values
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Set Price values
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWS As Worksheet
    Set pivotWS = Worksheets.Add
    
    ' Create the Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotWS.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)
    
    ' Add fields to the pivot table
    With pivotTable
        ' Set 'Style' as column field
        .PivotFields("Style").Orientation = xlColumnField
        ' Set 'Region' as row field
        .PivotFields("Region").Orientation = xlRowField
        ' Add 'Price' as data field with sum aggregation
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Set values in the pivot table worksheet
    pivotWS.Range("A13").Value = "Drag to column"
    pivotWS.Range("B13").Value = pivotTable.PivotFields("Region").DragToColumn
End Sub
```

## OnlyOffice JS Code

```javascript
// JavaScript code to manipulate worksheet and create a pivot table using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
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

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    columns: ['Style'],  // Set 'Style' as column field
    rows: 'Region',      // Set 'Region' as row field
});

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set values in the pivot table worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to column');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToColumn());
```