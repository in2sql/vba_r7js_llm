# Description / Описание

This code sets up data in an active worksheet, creates a pivot table from the data, configures the pivot table fields, and checks if the 'Region' field has repeat labels enabled.  
Этот код задает данные в активном листе, создает сводную таблицу из данных, настраивает поля сводной таблицы и проверяет, включены ли повторяющиеся метки для поля «Регион».

## JavaScript Code / JavaScript Код

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data for Region
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set data for Style
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set data for Price
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Get the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Set the layout of row fields to Tabular
pivotTable.SetRowAxisLayout('Tabular');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set label for repeat labels information
pivotWorksheet.GetRange('A12').SetValue('Region repeat labels');

// Set the value indicating whether repeat labels are enabled
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetRepeatLabels());
```

## VBA Code / VBA Код

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Set data for Region
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Set data for Style
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Set data for Price
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = Worksheets.Add
pivotSheet.Name = "PivotSheet"

' Create the pivot table
Dim pivotTable As PivotTable
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRef)
Set pivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotSheet.Range("A1"), _
    TableName:="PivotTable1")

' Add row fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Region").Position = 1
    .PivotFields("Style").Orientation = xlRowField
    .PivotFields("Style").Position = 2
End With

' Add data field to the pivot table
With pivotTable
    .PivotFields("Price").Orientation = xlDataField
    .PivotFields("Price").Function = xlSum
    .PivotFields("Price").Name = "Sum of Price"
End With

' Set the layout of row fields to Tabular
pivotSheet.PivotTables("PivotTable1").RowAxisLayout xlTabularRow

' Set label for repeat labels information
pivotSheet.Range("A12").Value = "Region repeat labels"

' Check if repeat labels are enabled for 'Region' field
Dim repeatLabels As Boolean
repeatLabels = pivotTable.PivotFields("Region").RepeatLabels

' Set the value indicating whether repeat labels are enabled
pivotSheet.Range("B12").Value = repeatLabels
```