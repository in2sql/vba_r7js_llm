**Creating worksheet data and a pivot table**  
**Создание данных листа и сводной таблицы**

```vba
' VBA Code
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim pivotWs As Worksheet
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache

    ' Get the active sheet
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

    ' Add a new worksheet for the pivot table
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
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Set a value in the pivot worksheet
    pivotWs.Range("A12").Value = "The Style field position will change soon"

    ' Change the position of the 'Style' field after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "SetPivotFieldPosition"
End Sub

Sub SetPivotFieldPosition()
    Dim pivotTable As PivotTable
    Set pivotTable = ThisWorkbook.Worksheets("PivotSheet").PivotTables("PivotTable1")
    ' Move 'Style' field to first position
    pivotTable.PivotFields("Style").Position = 1
End Sub
```

```javascript
// JavaScript Code
// Creating worksheet data and a pivot table
// Создание данных листа и сводной таблицы

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

// Insert pivot table in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Set a value in the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();
pivotWorksheet.GetRange('A12').SetValue('The Style field position will change soon');

// Get pivot field 'Style'
var pivotField = pivotTable.GetPivotFields('Style');

// After 5 seconds, change the position of 'Style' field
setTimeout(function () {
    pivotField.SetPosition(1);
}, 5000);
```