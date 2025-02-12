**Description / Описание:**

This code initializes data in an Excel worksheet, creates a pivot table based on the specified data range, adds fields to the pivot table, and sets values related to the pivot table's data fields.  
Этот код инициализирует данные в рабочем листе Excel, создает сводную таблицу на основе указанного диапазона данных, добавляет поля в сводную таблицу и устанавливает значения, связанные с полями данных сводной таблицы.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Set data for 'Region'
oWorksheet.GetRange('B2').SetValue('East');   // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');   // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');   // Set 'West' in cell B5

// Set data for 'Style'
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set 'Tee' in cell C5

// Set data for 'Price'
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
    rows: ['Region', 'Style'],
});

// Add data field 'Price' to the pivot table
var dataField = pivotTable.AddDataField('Price');
dataField.SetPosition(1); // Set position of the data field

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Set values related to the pivot table
pivotWorksheet.GetRange('A15').SetValue('Sum of Price2 position:'); // Set label in cell A15
pivotWorksheet.GetRange('B15').SetValue(dataField.GetPosition());    // Set position value in cell B15
```

```vba
' VBA Code Equivalent

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Set headers
    oWorksheet.Range("B1").Value = "Region" ' Set 'Region' in cell B1
    oWorksheet.Range("C1").Value = "Style"  ' Set 'Style' in cell C1
    oWorksheet.Range("D1").Value = "Price"  ' Set 'Price' in cell D1

    ' Set data for 'Region'
    oWorksheet.Range("B2").Value = "East"   ' Set 'East' in cell B2
    oWorksheet.Range("B3").Value = "West"   ' Set 'West' in cell B3
    oWorksheet.Range("B4").Value = "East"   ' Set 'East' in cell B4
    oWorksheet.Range("B5").Value = "West"   ' Set 'West' in cell B5

    ' Set data for 'Style'
    oWorksheet.Range("C2").Value = "Fancy"  ' Set 'Fancy' in cell C2
    oWorksheet.Range("C3").Value = "Fancy"  ' Set 'Fancy' in cell C3
    oWorksheet.Range("C4").Value = "Tee"    ' Set 'Tee' in cell C4
    oWorksheet.Range("C5").Value = "Tee"    ' Set 'Tee' in cell C5

    ' Set data for 'Price'
    oWorksheet.Range("D2").Value = 42.5     ' Set 42.5 in cell D2
    oWorksheet.Range("D3").Value = 35.2     ' Set 35.2 in cell D3
    oWorksheet.Range("D4").Value = 12.3     ' Set 12.3 in cell D4
    oWorksheet.Range("D5").Value = 24.8     ' Set 24.8 in cell D5

    ' Define the data range for the pivot table
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")

    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"

    ' Create the pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRef)

    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWorksheet.Range("A3"), _
        TableName:="PivotTable1")

    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Region").Position = 1
    pivotTable.PivotFields("Style").Orientation = xlRowField
    pivotTable.PivotFields("Style").Position = 2

    ' Add data field 'Price'
    Dim dataField As PivotField
    Set dataField = pivotTable.PivotFields("Price")
    dataField.Orientation = xlDataField
    dataField.Function = xlSum
    dataField.Position = 1

    ' Set values related to the pivot table
    pivotWorksheet.Range("A15").Value = "Sum of Price2 position:" ' Set label in cell A15
    pivotWorksheet.Range("B15").Value = dataField.Position      ' Set position value in cell B15
End Sub
```