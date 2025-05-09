### Description / Описание

This code creates and populates a worksheet with data, inserts a pivot table based on that data, and sets up fields and data positions in both English and Russian.

Этот код создает и заполняет рабочий лист данными, вставляет сводную таблицу на основе этих данных и настраивает поля и позиции данных на английском и русском языках.

### VBA Code / Код VBA

```vba
' VBA code to create and populate a worksheet, insert pivot table, and set data fields

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim pivotRange As Range
    Dim dataField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Populate Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Populate Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Populate Price data
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
    Set pt = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add Row fields
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add Data field
    Set dataField = pt.PivotFields("Price")
    dataField.Orientation = xlDataField
    dataField.Function = xlSum
    dataField.Position = 1
    
    ' Set values in pivot worksheet
    pivotWs.Range("A15").Value = "Sum of Price2 position:"
    pivotWs.Range("B15").Value = dataField.Position
End Sub
```

### OnlyOffice JS Code / Код OnlyOffice JS

```javascript
// JavaScript code to create and populate a worksheet, insert pivot table, and set data fields

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

// Insert a new worksheet with a pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add Data field to the pivot table
pivotTable.AddDataField('Price');
var dataField = pivotTable.AddDataField('Price');

// Set the position of the data field
dataField.SetPosition(1);

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A15').SetValue('Sum of Price2 position:');
pivotWorksheet.GetRange('B15').SetValue(dataField.GetPosition());
```