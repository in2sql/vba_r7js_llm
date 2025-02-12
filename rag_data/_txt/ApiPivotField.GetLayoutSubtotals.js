**VBA and OnlyOffice JS code to set cell values and create a pivot table.  
VBA и OnlyOffice JS код для установки значений ячеек и создания сводной таблицы.**

```vba
' VBA Code to set cell values and create a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
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
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the pivot field 'Region'
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set values related to pivot field
    pivotWs.Range("A14").Value = "Region layout subtotals"
    pivotWs.Range("B14").Value = pivotField.Subtotals(1) ' xlSubtotalAutomatic
End Sub
```

```javascript
// OnlyOffice JS Code to set cell values and create a pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set the header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set the header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set the header for Price

// Set data values
oWorksheet.GetRange('B2').SetValue('East');   // Set Region for row 2
oWorksheet.GetRange('B3').SetValue('West');   // Set Region for row 3
oWorksheet.GetRange('B4').SetValue('East');   // Set Region for row 4
oWorksheet.GetRange('B5').SetValue('West');   // Set Region for row 5

oWorksheet.GetRange('C2').SetValue('Fancy');  // Set Style for row 2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set Style for row 3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set Style for row 4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set Style for row 5

oWorksheet.GetRange('D2').SetValue(42.5);      // Set Price for row 2
oWorksheet.GetRange('D3').SetValue(35.2);      // Set Price for row 3
oWorksheet.GetRange('D4').SetValue(12.3);      // Set Price for row 4
oWorksheet.GetRange('D5').SetValue(24.8);      // Set Price for row 5

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Set Region and Style as row fields
});

// Add data field to the pivot table
pivotTable.AddDataField('Price'); // Add Price as a data field

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set values related to pivot field
pivotWorksheet.GetRange('A14').SetValue('Region layout subtotals'); // Set description
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetLayoutSubtotals()); // Set subtotal layout
```