**Description / Описание:**
This script populates an Excel worksheet with data, creates a pivot table from that data, and sets a value based on the pivot table's data field.  
Этот скрипт заполняет рабочий лист Excel данными, создает сводную таблицу на основе этих данных и устанавливает значение на основе поля данных сводной таблицы.

```vba
' VBA Code to replicate the OnlyOffice API functionality

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region" ' Set header for Region
    oWorksheet.Range("C1").Value = "Style"  ' Set header for Style
    oWorksheet.Range("D1").Value = "Price"  ' Set header for Price
    
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
    
    ' Define the data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the PivotTable
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A3"), TableName:="SalesPivot")
    
    ' Add Row Fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add Data Field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the data field name
    Dim dataFieldName As String
    dataFieldName = pivotTable.PivotFields("Sum of Price").Name
    
    ' Set values in the pivot sheet
    pivotSheet.Range("A12").Value = "The Data field name"
    pivotSheet.Range("B12").Value = dataFieldName
End Sub
```

```javascript
// JavaScript Code to replicate the OnlyOffice API functionality

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new worksheet with a pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'], // Add Region and Style as row fields
});

// Add data field
pivotTable.AddDataField('Price'); // Add Price as the data field

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field name
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in the pivot sheet
pivotWorksheet.GetRange('A12').SetValue('The Data field name'); // Label for data field name
pivotWorksheet.GetRange('B12').SetValue(dataField.GetName());    // Set the data field name
```