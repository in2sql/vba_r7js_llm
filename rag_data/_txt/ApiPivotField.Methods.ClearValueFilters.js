**Description:**
- **English:** This script populates the active worksheet with region, style, and price data. It then creates a pivot table in a new worksheet, setting 'Region' as rows, 'Style' as columns, and 'Price' as the data field. Finally, it clears any value filters applied to the 'Region' field in the pivot table.
- **Russian:** Этот скрипт заполняет активный лист данными о регионе, стиле и цене. Затем он создаёт сводную таблицу на новом листе, устанавливая 'Region' в качестве строк, 'Style' в качестве столбцов и 'Price' в качестве поля данных. В конце он очищает любые фильтры значений, применённые к полю 'Region' в сводной таблице.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style'); // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price'); // Set header for Price

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');  // Row 2 Region
oWorksheet.GetRange('B3').SetValue('West');  // Row 3 Region
oWorksheet.GetRange('B4').SetValue('East');  // Row 4 Region
oWorksheet.GetRange('B5').SetValue('West');  // Row 5 Region

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy'); // Row 2 Style
oWorksheet.GetRange('C3').SetValue('Fancy'); // Row 3 Style
oWorksheet.GetRange('C4').SetValue('Tee');   // Row 4 Style
oWorksheet.GetRange('C5').SetValue('Tee');   // Row 5 Style

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5); // Row 2 Price
oWorksheet.GetRange('D3').SetValue(35.2); // Row 3 Price
oWorksheet.GetRange('D4').SetValue(12.3); // Row 4 Price
oWorksheet.GetRange('D5').SetValue(24.8); // Row 5 Price

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
	rows: 'Region',  // Set 'Region' as row field
	columns: 'Style' // Set 'Style' as column field
});

// Add 'Price' as the data field in the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet containing the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' field in the pivot table
var pivotField = pivotTable.GetPivotFields('Region');

// Clear any value filters applied to the 'Region' field
pivotField.ClearValueFilters(); 
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region" ' Set header for Region
oWorksheet.Range("C1").Value = "Style"  ' Set header for Style
oWorksheet.Range("D1").Value = "Price"  ' Set header for Price

' Populate Region data
oWorksheet.Range("B2").Value = "East"  ' Row 2 Region
oWorksheet.Range("B3").Value = "West"  ' Row 3 Region
oWorksheet.Range("B4").Value = "East"  ' Row 4 Region
oWorksheet.Range("B5").Value = "West"  ' Row 5 Region

' Populate Style data
oWorksheet.Range("C2").Value = "Fancy" ' Row 2 Style
oWorksheet.Range("C3").Value = "Fancy" ' Row 3 Style
oWorksheet.Range("C4").Value = "Tee"   ' Row 4 Style
oWorksheet.Range("C5").Value = "Tee"   ' Row 5 Style

' Populate Price data
oWorksheet.Range("D2").Value = 42.5 ' Row 2 Price
oWorksheet.Range("D3").Value = 35.2 ' Row 3 Price
oWorksheet.Range("D4").Value = 12.3 ' Row 4 Price
oWorksheet.Range("D5").Value = 24.8 ' Row 5 Price

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Insert a new pivot table in a new worksheet
Dim pivotWorkbook As Workbook
Dim pivotWorksheet As Worksheet
Dim pivotTable As PivotTable
Dim pivotCache As PivotCache

Set pivotWorkbook = ThisWorkbook
Set pivotWorksheet = pivotWorkbook.Worksheets.Add
Set pivotCache = pivotWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A3"), TableName:="PivotTable1")

' Add fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField ' Set 'Region' as row field
    .PivotFields("Style").Orientation = xlColumnField ' Set 'Style' as column field
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Set 'Price' as data field
End With

' Clear any value filters applied to the 'Region' field
With pivotTable.PivotFields("Region")
    .ClearAllFilters
End With
```