**Description:**

English: This code populates an Excel worksheet with sample data, creates a pivot table based on this data, and modifies the pivot table's field orientation after a delay.

Russian: Этот код заполняет рабочий лист Excel примерными данными, создает на основе этих данных сводную таблицу и изменяет ориентацию поля сводной таблицы после задержки.

---

**VBA Code:**

```vba
' Populate data, create pivot table, and modify field orientation after delay

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    
    ' Set the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data values
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
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Add a message to the pivot worksheet
    pivotWs.Range("A12").Value = "The Style field orientation will change soon"
    
    ' Get the pivot field 'Style'
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Change orientation after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "ChangePivotFieldOrientation"
End Sub

Sub ChangePivotFieldOrientation()
    Dim pt As PivotTable
    Dim pf As PivotField
    
    ' Assuming the pivot table is named "PivotTable1" and is on "PivotSheet"
    Set pt = ThisWorkbook.Worksheets("PivotSheet").PivotTables("PivotTable1")
    Set pf = pt.PivotFields("Style")
    
    ' Set orientation to column
    pf.Orientation = xlColumnField
End Sub
```

---

**OnlyOffice JavaScript Code:**

```javascript
// Populate data, create pivot table, and modify field orientation after delay

var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data values
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

// Insert pivot table into a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Add a message to the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();
pivotWorksheet.GetRange('A12').SetValue('The Style field orientation will change soon');

// Get the pivot field 'Style'
var pivotField = pivotTable.GetPivotFields('Style');

// Change orientation to columns after 5 seconds
setTimeout(function () {
    pivotField.SetOrientation("Columns");
}, 5000);
```