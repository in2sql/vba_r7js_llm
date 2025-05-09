**Description:**
This script populates an Excel sheet with data, creates a pivot table, and sets data field captions.
**Описание:**
Этот скрипт заполняет лист Excel данными, создает сводную таблицу и задает заголовки полей данных.

```vba
' Excel VBA code

Sub CreatePivotTable()
    ' Get the active worksheet
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
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pCache As PivotCache
    Set pCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    Dim pTable As PivotTable
    Set pTable = pCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add row fields
    pTable.PivotFields("Region").Orientation = xlRowField
    pTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pTable.AddDataField pTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set data field caption
    pivotWs.Range("A12").Value = "The Data field caption"
    pivotWs.Range("B12").Value = pTable.PivotFields("Sum of Price").Caption
End Sub
```

```javascript
// OnlyOffice JavaScript code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set data field caption
pivotWorksheet.GetRange('A12').SetValue('The Data field caption');
pivotWorksheet.GetRange('B12').SetValue(dataField.GetCaption());
```