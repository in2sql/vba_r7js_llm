**Description:**
This code sets up data in an Excel worksheet, creates a pivot table based on the data, and modifies the data field name.
Этот код устанавливает данные в рабочем листе Excel, создает сводную таблицу на основе данных и изменяет имя поля данных.

```vba
' VBA Code to replicate the OnlyOffice JS functionality
Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data
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
    
    ' Define the data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = Sheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Modify data field name
    Dim dataField As PivotField
    Set dataField = pivotTable.PivotFields("Sum of Price")
    
    ' Set values in Pivot worksheet
    pivotWorksheet.Range("A12").Value = "Data field name"
    pivotWorksheet.Range("B12").Value = dataField.Name
    
    dataField.Name = "My Sum of Price"
    pivotWorksheet.Range("A13").Value = "New Data field name"
    pivotWorksheet.Range("B13").Value = dataField.Name
End Sub
```

```javascript
// OnlyOffice JS code to set up data and create a pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the active sheet for the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Sum of Price' data field
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Data field name');
pivotWorksheet.GetRange('B12').SetValue(dataField.GetName());

// Rename the data field
dataField.SetName('My Sum of Price');
pivotWorksheet.GetRange('A13').SetValue('New Data field name');
pivotWorksheet.GetRange('B13').SetValue(dataField.GetName());
```