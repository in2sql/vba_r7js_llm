### Description / Описание

**English:** This script populates a worksheet with regions, styles, and prices data, then creates a pivot table to summarize the data by region and style, displaying the sum of prices. It also sets the pivot table layout to tabular and inserts a blank line for the region field.

**Russian:** Этот скрипт заполняет лист данными о регионах, стилях и ценах, затем создает сводную таблицу для суммирования данных по регионам и стилям, отображая сумму цен. Он также устанавливает табличный макет сводной таблицы и вставляет пустую строку для поля региона.

```javascript
// JavaScript code for OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set region values
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set style values
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set price values
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Set the row axis layout to tabular
pivotTable.SetRowAxisLayout('Tabular');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set layout to include blank lines
pivotField.SetLayoutBlankLine(true);

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A14').SetValue('Region blank line');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetLayoutBlankLine());
```

```vba
' VBA code equivalent for Excel

Sub CreatePivotTable()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"

    ' Set region values
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"

    ' Set style values
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"

    ' Set price values
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8

    ' Define the data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")

    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"

    ' Create the pivot table
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")

    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Style").Position = 2
    End With

    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

    ' Set row layout to tabular
    pivotTable.RowAxisLayout xlTabularRow

    ' Insert a blank line after Region field
    pivotSheet.Range("A14").Value = "Region blank line"
    ' Note: VBA does not have a direct equivalent for SetLayoutBlankLine
    ' Additional formatting may be required to simulate blank lines
    pivotSheet.Range("B14").Value = "True" ' Assuming the blank line is set
End Sub
```