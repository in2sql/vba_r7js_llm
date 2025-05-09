## Description / Описание

**English:**  
The following code sets up a worksheet by populating specific cells with data, creates a pivot table based on that data, and then modifies the pivot table by removing a data field after a short delay.

**Russian:**  
Следующий код настраивает рабочий лист, заполняя определенные ячейки данными, создает сводную таблицу на основе этих данных, а затем модифицирует сводную таблицу, удаляя поле данных после небольшой задержки.

```vba
' VBA Code to replicate the OnlyOffice JS functionality

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate header cells
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
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add Row Fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add Data Field
    Set dataField = pivotTable.PivotFields("Price")
    dataField.Orientation = xlDataField
    dataField.Function = xlSum
    dataField.Name = "Sum of Price"
    
    ' Add a message to the pivot sheet
    pivotWs.Range("A12").Value = "Sum of Price will be deleted soon"
    
    ' Schedule the removal of the data field after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "RemoveDataField"
End Sub

Sub RemoveDataField()
    Dim pivotWs As Worksheet
    Dim pivotTable As PivotTable
    
    ' Set the pivot table worksheet
    Set pivotWs = ThisWorkbook.Worksheets("PivotTableSheet")
    
    ' Set the pivot table
    Set pivotTable = pivotWs.PivotTables("SalesPivotTable")
    
    ' Remove the data field
    On Error Resume Next
    pivotTable.PivotFields("Sum of Price").Orientation = xlHidden
    On Error GoTo 0
End Sub
```

```javascript
// OnlyOffice JS Code with English comments

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

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Row Fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add a Data Field to the pivot table
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field from the pivot table
var dataField = pivotTable.GetDataFields('Sum of Price');

// Add a message to the pivot sheet
pivotWorksheet.GetRange('A12').SetValue('Sum of Price will be deleted soon');

// Schedule the removal of the data field after 5 seconds
setTimeout(function() {
	dataField.Remove();
}, 5000);
```