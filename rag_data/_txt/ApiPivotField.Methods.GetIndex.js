### Description / Описание

**English:** This code populates specific cells with data, creates a pivot table based on the data, and retrieves the index of the 'Style' pivot field.

**Русский:** Этот код заполняет определенные ячейки данными, создает сводную таблицу на основе данных и получает индекс поля сводной таблицы 'Style'.

```vba
' VBA Code to populate cells, create a pivot table, and get the index of the "Style" field
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim pivotWs As Worksheet
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    
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
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Set pivotTable = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the 'Style' pivot field and its index
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Display the index of the 'Style' field
    pivotWs.Range("A12").Value = "The Style field index"
    pivotWs.Range("B12").Value = pivotField.Position
End Sub
```

```javascript
// OnlyOffice JS Code to populate cells, create a pivot table, and get the index of the "Style" field

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

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add 'Price' as a data field in the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet where the pivot table is created
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set values to display the index of the 'Style' field
pivotWorksheet.GetRange('A12').SetValue('The Style field index');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetIndex());
```