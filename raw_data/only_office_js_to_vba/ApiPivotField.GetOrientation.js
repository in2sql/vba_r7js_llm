**Description:**
This code sets up data in an OnlyOffice worksheet, creates a pivot table, and retrieves field orientation.
Этот код настраивает данные в рабочем листе OnlyOffice, создает сводную таблицу и получает ориентацию поля.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set price data
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
	rows: 'Region',
	columns: 'Style',
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set descriptions in pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('The Style field orientation');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetOrientation());
```

```vba
' This VBA code sets up data in the active worksheet, creates a pivot table, and retrieves field orientation.
' Этот VBA-код настраивает данные в активном листе, создает сводную таблицу и получает ориентацию поля.

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set price data
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
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Set descriptions in pivot worksheet
    pivotWs.Range("A12").Value = "The Style field orientation"
    pivotWs.Range("B12").Value = pivotTable.PivotFields("Style").Orientation
End Sub
```