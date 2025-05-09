**Description:** This code populates specific cells in a worksheet with data and creates a pivot table based on that data.  
**Описание:** Этот код заполняет определенные ячейки на листе данными и создает сводную таблицу на основе этих данных.

```javascript
// OnlyOffice JavaScript code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set Region values
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set Style values
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set Price values
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Reference the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to the pivot table
pivotTable.AddFields({
	rows: 'Region',
	columns: 'Style',
});

// Get the active sheet for the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Add 'Price' as a data field in the pivot table
pivotTable.AddDataField('Price');

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set a value in the pivot worksheet
pivotWorksheet.GetRange('A10').SetValue('The Region field will be moved soon');

// After 5 seconds, move the 'Region' field to columns
setTimeout(function () {
	pivotField.Move('Columns');
}, 5000);
```

```vba
' Excel VBA code

' Description: This code populates specific cells in a worksheet with data and creates a pivot table based on that data.
' Описание: Этот код заполняет определенные ячейки на листе данными и создает сводную таблицу на основе этих данных.

Sub CreatePivotTable()
    ' Declare variables
    Dim oWorksheet As Worksheet
    Dim pivotWorksheet As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField

    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"

    ' Set Region values
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"

    ' Set Style values
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"

    ' Set Price values
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8

    ' Set the data range
    Set dataRange = oWorksheet.Range("B1:D5")

    ' Create a new worksheet for the pivot table
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"

    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWorksheet.Range("A1"), _
        TableName:="SalesPivotTable")

    ' Add 'Region' to Rows
    With pivotTable.PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add 'Style' to Columns
    With pivotTable.PivotFields("Style")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' Add 'Price' to Values
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Price"
    End With

    ' Set a value in the pivot worksheet
    pivotWorksheet.Range("A10").Value = "The Region field will be moved soon"

    ' Move 'Region' field to Columns after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "MoveRegionField"
End Sub

Sub MoveRegionField()
    ' Declare variables
    Dim pivotWorksheet As Worksheet
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField

    ' Set the pivot worksheet
    Set pivotWorksheet = ThisWorkbook.Worksheets("PivotTableSheet")

    ' Set the pivot table
    Set pivotTable = pivotWorksheet.PivotTables("SalesPivotTable")

    ' Get the 'Region' pivot field
    Set pivotField = pivotTable.PivotFields("Region")

    ' Move 'Region' to Columns
    With pivotField
        .Orientation = xlColumnField
        .Position = 2
    End With

    ' Update the cell value
    pivotWorksheet.Range("A10").Value = "The Region field has been moved to Columns"
End Sub
```