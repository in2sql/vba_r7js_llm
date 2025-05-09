**Description / Описание**

This script sets up a dataset with regions, styles, and prices, creates a pivot table based on this data, and manipulates a specific pivot field to demonstrate its functionality.

Этот скрипт устанавливает набор данных с регионами, стилями и ценами, создает сводную таблицу на основе этих данных и манипулирует определенным полем сводной таблицы для демонстрации его функциональности.

---

### OnlyOffice JS Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate Region column
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate Style column
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price column
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Set the layout of the row axis to Tabular form
pivotTable.SetRowAxisLayout("Tabular", false);

// Add the Price field as a data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' field in the pivot table
var pivotField = pivotTable.GetPivotFields('Style');

// Set values and manipulate the 'Style' field
pivotWorksheet.GetRange('A12').SetValue('Style field value');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetValue());

pivotWorksheet.GetRange('A14').SetValue('New Style field value');
pivotField.SetValue('My value');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetValue());
```

---

### Excel VBA Code

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Populate Region column
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Populate Style column
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Populate Price column
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define data range for pivot table
Dim dataRange As Range
Set dataRange = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add
pivotSheet.Name = "PivotSheet"

' Create the Pivot Cache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

' Create the Pivot Table
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotSheet.Range("A1"), _
    TableName:="SalesPivotTable")

' Add fields to pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
    .PivotFields("Price").Orientation = xlDataField
    .PivotFields("Price").Function = xlSum
End With

' Set the layout of the row axis to Tabular form
pivotTable.RowAxisLayout xlTabularRow

' Manipulate the 'Style' field
Dim pivotField As PivotField
Set pivotField = pivotTable.PivotFields("Style")

' Set values in the pivot sheet
With pivotSheet
    .Range("A12").Value = "Style field value"
    .Range("B12").Value = pivotField.CurrentPage
    
    .Range("A14").Value = "New Style field value"
    pivotField.CurrentPage = "My value"
    .Range("B14").Value = pivotField.CurrentPage
End With
```