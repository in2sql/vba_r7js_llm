**Description / Описание**

English:
This code sets up data in a worksheet, creates a pivot table, adds fields, and sets captions.

Russian:
Этот код заполняет данные в листе, создает сводную таблицу, добавляет поля и устанавливает заголовки.

```vba
' VBA Code to perform the equivalent operations

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in cells
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
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
    
    ' Create a pivot table
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Set caption for the Style field
    pivotSheet.Range("A12").Value = "The Style field caption"
    pivotSheet.Range("B12").Value = pivotTable.PivotFields("Style").Caption
End Sub
```

```javascript
// OnlyOffice JS Code to perform the equivalent operations

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

oWorksheet.GetRange('B2').SetValue('East');    // Set data
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

oWorksheet.GetRange('C2').SetValue('Fancy');   // Set data
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

oWorksheet.GetRange('D2').SetValue(42.5);      // Set data
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Add Region and Style as row fields
});

// Add data field
pivotTable.AddDataField('Price'); // Add Price as data field

// Get the new pivot worksheet and pivot field
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Style');

// Set caption for the Style field
pivotWorksheet.GetRange('A12').SetValue('The Style field caption');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetCaption());
```