**Description / Описание**

This code initializes an active worksheet, sets values in specific cells, creates a pivot table based on a data range, adds fields to the pivot table, configures its layout, and retrieves a property from the pivot table.  
Этот код инициализирует активный рабочий лист, устанавливает значения в определенные ячейки, создает сводную таблицу на основе диапазона данных, добавляет поля в сводную таблицу, настраивает ее макет и извлекает свойство из сводной таблицы.

```javascript
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

// Get the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Set the row axis layout to Tabular
pivotTable.SetRowAxisLayout('Tabular');

// Get the active worksheet containing the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' field from the pivot table
var pivotField = pivotTable.GetPivotFields('Region');

// Set a value in the pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Region repeat labels');

// Get and set whether the 'Region' labels are repeated
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetRepeatLabels());
```

```vba
' Initialize the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

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

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = Worksheets.Add

' Create the pivot table
Dim pivotTable As PivotTable
Set pivotTable = pivotSheet.PivotTables.Add(PivotCache:=ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, SourceData:=dataRef), TableDestination:=pivotSheet.Range("A3"), TableName:="PivotTable1")

' Add row fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
End With

' Add data field to the pivot table
pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

' Set the row axis layout to Tabular
pivotTable.RowAxisLayout xlTabularRow

' Get the 'Region' field from the pivot table
Dim pivotField As PivotField
Set pivotField = pivotTable.PivotFields("Region")

' Set a value in the pivot worksheet
pivotSheet.Range("A12").Value = "Region repeat labels"

' Get and set whether the 'Region' labels are repeated
pivotSheet.Range("B12").Value = pivotField.RepeatLabels

```