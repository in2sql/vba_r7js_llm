# Description / Описание

This code populates data in a worksheet, creates a pivot table, and retrieves a data field value in OnlyOffice API and Excel VBA.

Этот код заполняет данные в листе, создает сводную таблицу и извлекает значение поля данных в OnlyOffice API и Excel VBA.

```vba
' VBA code

Sub CreatePivotTable()
    ' Set the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set header values
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

    ' Define data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")

    ' Create a new worksheet for the Pivot Table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"

    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)

    ' Add fields to Pivot Table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Get data field value
    Dim dataFieldValue As Double
    dataFieldValue = pivotTable.GetPivotData("Sum of Price").Value

    ' Set value in pivot worksheet
    pivotSheet.Range("A12").Value = "The Data field value"
    pivotSheet.Range("B12").Value = dataFieldValue
End Sub
```

```javascript
// OnlyOffice JS code

// Get the active worksheet
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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert Pivot Table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to Pivot Table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field to Pivot Table
pivotTable.AddDataField('Price');

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Retrieve the data field value
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('The Data field value');
pivotWorksheet.GetRange('B12').SetValue(dataField.GetValue());
```