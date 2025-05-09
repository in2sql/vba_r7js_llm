**Description / Описание:**

English: The following VBA and OnlyOffice JavaScript code initializes a worksheet by setting values in specific cells, creates a pivot table based on a data range, adds row and data fields to the pivot table, and retrieves the index of a specific pivot field.

Russian: Следующий код на VBA и OnlyOffice JavaScript инициализирует рабочий лист, устанавливая значения в определенные ячейки, создает сводную таблицу на основе диапазона данных, добавляет поля строк и данных в сводную таблицу и получает индекс конкретного поля сводной таблицы.

```vba
' VBA code to manipulate worksheet and create pivot table

Sub CreatePivotTable()

    ' Set references to the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in header row
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Populate Region data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Populate Style data
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Populate Price data
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define data range for pivot table
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    
    ' Create the pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A3"), TableName:="PivotTable1")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.PivotFields("Price").Orientation = xlDataField
    pivotTable.PivotFields("Price").Function = xlSum
    
    ' Retrieve the index of the 'Style' field
    Dim pivotFieldIndex As Integer
    pivotFieldIndex = pivotTable.PivotFields("Style").Position
    
    ' Set the value in pivot sheet
    pivotSheet.Range("A12").Value = "The Style field index"
    pivotSheet.Range("B12").Value = pivotFieldIndex

End Sub
```

```js
// OnlyOffice JavaScript code to manipulate worksheet and create pivot table

// Initialize the API and get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Set values in header row
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set Region to East
oWorksheet.GetRange('B3').SetValue('West');   // Set Region to West
oWorksheet.GetRange('B4').SetValue('East');   // Set Region to East
oWorksheet.GetRange('B5').SetValue('West');   // Set Region to West

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set Style to Fancy
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set Style to Fancy
oWorksheet.GetRange('C4').SetValue('Tee');    // Set Style to Tee
oWorksheet.GetRange('C5').SetValue('Tee');    // Set Style to Tee

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);      // Set Price
oWorksheet.GetRange('D3').SetValue(35.2);      // Set Price
oWorksheet.GetRange('D4').SetValue(12.3);      // Set Price
oWorksheet.GetRange('D5').SetValue(24.8);      // Set Price

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new worksheet with the pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Add Region and Style as row fields
});

// Add Price as a data field
pivotTable.AddDataField('Price'); // Add Price as data field

// Get the active sheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field 'Style'
var pivotField = pivotTable.GetPivotFields('Style');

// Set description and the index of the 'Style' field
pivotWorksheet.GetRange('A12').SetValue('The Style field index'); // Description
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetIndex());    // Set the index value
```