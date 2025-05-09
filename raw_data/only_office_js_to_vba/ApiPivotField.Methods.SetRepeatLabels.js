**Description:**  
This script populates an Excel worksheet with sample data, creates a pivot table based on that data, and configures the pivot table's layout and fields.  
Этот скрипт заполняет рабочий лист Excel примерными данными, создает сводную таблицу на основе этих данных и настраивает макет и поля сводной таблицы.

```vba
' VBA Code Equivalent to OnlyOffice API Script

Sub CreatePivotTable()
    ' Define the worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
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
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Sheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add Row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add Data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set row layout to Tabular
    pivotSheet.PivotTables("SalesPivotTable").RowAxisLayout xlTabularRow
    
    ' Repeat Labels for Region
    pivotSheet.PivotTables("SalesPivotTable").PivotFields("Region").RepeatLabels = True
    
    ' Set values in A12 and B12
    pivotSheet.Range("A12").Value = "Region repeat labels"
    pivotSheet.Range("B12").Value = pivotSheet.PivotTables("SalesPivotTable").PivotFields("Region").RepeatLabels
End Sub
```

```javascript
// OnlyOffice JS Code

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

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add Data field
pivotTable.AddDataField('Price');

// Set row layout to Tabular
pivotTable.SetRowAxisLayout('Tabular');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field and set repeat labels
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.SetRepeatLabels(true);

// Set values in A12 and B12
pivotWorksheet.GetRange('A12').SetValue('Region repeat labels');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetRepeatLabels());
```