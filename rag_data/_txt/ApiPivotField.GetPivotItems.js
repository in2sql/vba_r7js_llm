# Create and Populate a Worksheet with Data and a Pivot Table
# Создание и заполнение листа данными и сводной таблицей

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set cell B1 to 'Region'
oWorksheet.GetRange('C1').SetValue('Style');  // Set cell C1 to 'Style'
oWorksheet.GetRange('D1').SetValue('Price');  // Set cell D1 to 'Price'

// Populate 'Region' column
oWorksheet.GetRange('B2').SetValue('East');   // Set cell B2 to 'East'
oWorksheet.GetRange('B3').SetValue('West');   // Set cell B3 to 'West'
oWorksheet.GetRange('B4').SetValue('East');   // Set cell B4 to 'East'
oWorksheet.GetRange('B5').SetValue('West');   // Set cell B5 to 'West'

// Populate 'Style' column
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set cell C2 to 'Fancy'
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set cell C3 to 'Fancy'
oWorksheet.GetRange('C4').SetValue('Tee');    // Set cell C4 to 'Tee'
oWorksheet.GetRange('C5').SetValue('Tee');    // Set cell C5 to 'Tee'

// Populate 'Price' column
oWorksheet.GetRange('D2').SetValue(42.5);     // Set cell D2 to 42.5
oWorksheet.GetRange('D3').SetValue(35.2);     // Set cell D3 to 35.2
oWorksheet.GetRange('D4').SetValue(12.3);     // Set cell D4 to 12.3
oWorksheet.GetRange('D5').SetValue(24.8);     // Set cell D5 to 24.8

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
	rows: 'Region',    // Set 'Region' as row field
	columns: 'Style',  // Set 'Style' as column field
});

// Get the active sheet where the pivot table is placed
var pivotWorksheet = Api.GetActiveSheet();

// Add 'Price' as the data field in the pivot table
pivotTable.AddDataField('Price');

// Get the 'Region' pivot field and its items
var pivotField = pivotTable.GetPivotFields('Region');
var pivotItems = pivotField.GetPivotItems();

// Set the header for pivot items
pivotWorksheet.GetRange('A10').SetValue('Region pivot items'); // Set cell A10 to 'Region pivot items'

// Populate the pivot items below the header
for (var i = 0; i < pivotItems.length; i += 1) {
	pivotWorksheet.GetRangeByNumber(9 + i, 1).SetValue(pivotItems[i].GetName()); // Set each pivot item name
} 
```

```vba
' Create and Populate a Worksheet with Data and a Pivot Table
' Создание и заполнение листа данными и сводной таблицей

Sub CreatePivotTable()
    ' Declare variables
    Dim ws As Worksheet
    Dim ptSheet As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    Dim pivotItem As PivotItem
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"   ' Set cell B1 to 'Region'
    ws.Range("C1").Value = "Style"    ' Set cell C1 to 'Style'
    ws.Range("D1").Value = "Price"    ' Set cell D1 to 'Price'
    
    ' Populate 'Region' column
    ws.Range("B2").Value = "East"     ' Set cell B2 to 'East'
    ws.Range("B3").Value = "West"     ' Set cell B3 to 'West'
    ws.Range("B4").Value = "East"     ' Set cell B4 to 'East'
    ws.Range("B5").Value = "West"     ' Set cell B5 to 'West'
    
    ' Populate 'Style' column
    ws.Range("C2").Value = "Fancy"    ' Set cell C2 to 'Fancy'
    ws.Range("C3").Value = "Fancy"    ' Set cell C3 to 'Fancy'
    ws.Range("C4").Value = "Tee"      ' Set cell C4 to 'Tee'
    ws.Range("C5").Value = "Tee"      ' Set cell C5 to 'Tee'
    
    ' Populate 'Price' column
    ws.Range("D2").Value = 42.5       ' Set cell D2 to 42.5
    ws.Range("D3").Value = 35.2       ' Set cell D3 to 35.2
    ws.Range("D4").Value = 12.3       ' Set cell D4 to 12.3
    ws.Range("D5").Value = 24.8       ' Set cell D5 to 24.8
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set ptSheet = ThisWorkbook.Worksheets.Add
    ptSheet.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=ptSheet.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add 'Region' to rows
    pivotTable.PivotFields("Region").Orientation = xlRowField  ' Set 'Region' as row field
    
    ' Add 'Style' to columns
    pivotTable.PivotFields("Style").Orientation = xlColumnField  ' Set 'Style' as column field
    
    ' Add 'Price' to values
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum  ' Set 'Price' as data field
    
    ' Add header for pivot items
    ptSheet.Range("A10").Value = "Region pivot items"  ' Set cell A10 to 'Region pivot items'
    
    ' Get the 'Region' pivot field
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Populate pivot items below the header
    i = 0
    For Each pivotItem In pivotField.PivotItems
        ptSheet.Cells(10 + i + 1, 1).Value = pivotItem.Name  ' Set each pivot item name
        i = i + 1
    Next pivotItem
End Sub
```