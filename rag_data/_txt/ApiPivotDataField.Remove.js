**English:** This code initializes a worksheet with specific data, creates a pivot table based on that data, adds fields to the pivot table, and removes a data field after a delay.

**Russian:** Этот код инициализирует лист с определенными данными, создает сводную таблицу на основе этих данных, добавляет поля в сводную таблицу и удаляет поле данных после задержки.

```javascript
// JavaScript Code for OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data for Region
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set data for Style
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set data for Price
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add a data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field 'Sum of Price'
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set a value in the pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Sum of Price will be deleted soon');

// Remove the data field after 5 seconds
setTimeout(function() {
    dataField.Remove();
}, 5000);
```

```vba
' VBA Code Equivalent

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"

    ' Set data for Region
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"

    ' Set data for Style
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"

    ' Set data for Price
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8

    ' Define the data range for the pivot table
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")

    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"

    ' Create the pivot table
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")

    ' Add row fields to the pivot table
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField

    ' Add data field to the pivot table
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

    ' Set a value in the pivot worksheet
    pivotSheet.Range("A12").Value = "Sum of Price will be deleted soon"

    ' Remove the data field after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "RemoveDataField"
End Sub

Sub RemoveDataField()
    ' Get the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets("PivotTableSheet")
    
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTables("PivotTable1")
    
    ' Remove the data field 'Sum of Price'
    pivotTable.PivotFields("Sum of Price").Orientation = xlHidden
    
    ' Update the pivot table
    pivotTable.RefreshTable
End Sub
```