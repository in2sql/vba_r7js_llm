**Description / Описание**

This script sets up a worksheet with region, style, and price data, creates a pivot table based on this data, and adds subtotals for the regions.  
Этот скрипт настраивает рабочий лист с данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных и добавляет промежуточные итоги по регионам.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Set region values
oWorksheet.GetRange('B2').SetValue('East');   // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');   // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');   // Set 'West' in cell B5

// Set style values
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set 'Tee' in cell C5

// Set price values
oWorksheet.GetRange('D2').SetValue(42.5);     // Set 42.5 in cell D2
oWorksheet.GetRange('D3').SetValue(35.2);     // Set 35.2 in cell D3
oWorksheet.GetRange('D4').SetValue(12.3);     // Set 12.3 in cell D4
oWorksheet.GetRange('D5').SetValue(24.8);     // Set 24.8 in cell D5

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
	columns: ['Style'], // Add 'Style' as column field
	rows: 'Region',     // Add 'Region' as row field
});

// Add data field to pivot table
pivotTable.AddDataField('Price'); // Add 'Price' as data field

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();
// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set subtotals for the 'Region' field
pivotField.SetSubtotals({
	Count: true, // Enable count subtotal
});

// Retrieve subtotals for the 'Region' field
var subtotals = pivotField.GetSubtotals();
// Set label for subtotals
pivotWorksheet.GetRange('A11').SetValue('Region subtotals');
// Populate subtotals starting from row 12
let k = 12;
for (var i in subtotals) {
	pivotWorksheet.GetRangeByNumber(k, 0).SetValue(i);       // Set region name
	pivotWorksheet.GetRangeByNumber(k++, 1).SetValue(subtotals[i]); // Set subtotal count
}
```

```vba
' VBA Code equivalent

Sub CreatePivotTableWithSubtotals()
    ' Set headers
    With ThisWorkbook.ActiveSheet
        .Range("B1").Value = "Region" ' Set 'Region' in cell B1
        .Range("C1").Value = "Style"  ' Set 'Style' in cell C1
        .Range("D1").Value = "Price"  ' Set 'Price' in cell D1
        
        ' Set region values
        .Range("B2").Value = "East"    ' Set 'East' in cell B2
        .Range("B3").Value = "West"    ' Set 'West' in cell B3
        .Range("B4").Value = "East"    ' Set 'East' in cell B4
        .Range("B5").Value = "West"    ' Set 'West' in cell B5
        
        ' Set style values
        .Range("C2").Value = "Fancy"   ' Set 'Fancy' in cell C2
        .Range("C3").Value = "Fancy"   ' Set 'Fancy' in cell C3
        .Range("C4").Value = "Tee"     ' Set 'Tee' in cell C4
        .Range("C5").Value = "Tee"     ' Set 'Tee' in cell C5
        
        ' Set price values
        .Range("D2").Value = 42.5      ' Set 42.5 in cell D2
        .Range("D3").Value = 35.2      ' Set 35.2 in cell D3
        .Range("D4").Value = 12.3      ' Set 12.3 in cell D4
        .Range("D5").Value = 24.8      ' Set 24.8 in cell D5
    End With
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = ThisWorkbook.ActiveSheet.Range("B1:D5")
    
    ' Add a new worksheet for pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create PivotCache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create PivotTable
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add 'Style' to Columns
    pivotTable.PivotFields("Style").Orientation = xlColumnField
    pivotTable.PivotFields("Style").Position = 1
    
    ' Add 'Region' to Rows
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Region").Position = 1
    
    ' Add 'Price' to Values
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Price"
    End With
    
    ' Enable count subtotal for 'Region'
    With pivotTable.PivotFields("Region")
        .Subtotals(1) = True ' Subtotals by Count
    End With
    
    ' Retrieve subtotals
    Dim subtotalRange As Range
    Set subtotalRange = pivotSheet.Range("A11")
    subtotalRange.Value = "Region subtotals" ' Set label for subtotals
    
    Dim i As Integer
    i = 12
    Dim pf As PivotField
    Set pf = pivotTable.PivotFields("Region")
    
    Dim item As PivotItem
    For Each item In pf.PivotItems
        pivotSheet.Cells(i, 1).Value = item.Name ' Set region name
        pivotSheet.Cells(i, 2).Value = item.DataRange.Cells(1, 1).Value ' Set subtotal count
        i = i + 1
    Next item
End Sub
```