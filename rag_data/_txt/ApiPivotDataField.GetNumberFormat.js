**Description / Описание:**
This code sets up data in a worksheet, creates a pivot table, and formats it accordingly.
Этот код заполняет данные в листе, создает сводную таблицу и выполняет ее форматирование.

```javascript
// JavaScript OnlyOffice API code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Set data for Region
oWorksheet.GetRange('B2').SetValue('East');  // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');  // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');  // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');  // Set 'West' in cell B5

// Set data for Style
oWorksheet.GetRange('C2').SetValue('Fancy'); // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy'); // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set 'Tee' in cell C5

// Set data for Price
oWorksheet.GetRange('D2').SetValue(42.5);     // Set 42.5 in cell D2
oWorksheet.GetRange('D3').SetValue(35.2);     // Set 35.2 in cell D3
oWorksheet.GetRange('D4').SetValue(12.3);     // Set 12.3 in cell D4
oWorksheet.GetRange('D5').SetValue(24.8);     // Set 24.8 in cell D5

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Add 'Region' and 'Style' as row fields
});

// Add data field to the pivot table
pivotTable.AddDataField('Price'); // Add 'Price' as a data field

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field for formatting
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set number format information in cells
pivotWorksheet.GetRange('A15').SetValue('Number format:');                     // Set label in A15
pivotWorksheet.GetRange('B15').SetValue(dataField.GetNumberFormat());          // Set number format in B15
```

```vba
' Excel VBA equivalent code

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region" ' Set 'Region' in cell B1
    oWorksheet.Range("C1").Value = "Style"  ' Set 'Style' in cell C1
    oWorksheet.Range("D1").Value = "Price"  ' Set 'Price' in cell D1
    
    ' Set data for Region
    oWorksheet.Range("B2").Value = "East"    ' Set 'East' in cell B2
    oWorksheet.Range("B3").Value = "West"    ' Set 'West' in cell B3
    oWorksheet.Range("B4").Value = "East"    ' Set 'East' in cell B4
    oWorksheet.Range("B5").Value = "West"    ' Set 'West' in cell B5
    
    ' Set data for Style
    oWorksheet.Range("C2").Value = "Fancy"   ' Set 'Fancy' in cell C2
    oWorksheet.Range("C3").Value = "Fancy"   ' Set 'Fancy' in cell C3
    oWorksheet.Range("C4").Value = "Tee"     ' Set 'Tee' in cell C4
    oWorksheet.Range("C5").Value = "Tee"     ' Set 'Tee' in cell C5
    
    ' Set data for Price
    oWorksheet.Range("D2").Value = 42.5      ' Set 42.5 in cell D2
    oWorksheet.Range("D3").Value = 35.2      ' Set 35.2 in cell D3
    oWorksheet.Range("D4").Value = 12.3      ' Set 12.3 in cell D4
    oWorksheet.Range("D5").Value = 24.8      ' Set 24.8 in cell D5
    
    ' Define the data range for the pivot table
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField ' Add 'Region' as row field
    pivotTable.PivotFields("Style").Orientation = xlRowField  ' Add 'Style' as row field
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum ' Add 'Price' as data field
    
    ' Set number format information in cells
    pivotSheet.Range("A15").Value = "Number format:"  ' Set label in A15
    pivotSheet.Range("B15").Value = pivotTable.PivotFields("Sum of Price").NumberFormat ' Set number format in B15
End Sub
```