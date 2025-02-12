### Description / Описание
**English:** This script populates an Excel sheet with region, style, and price data, creates a pivot table based on this data, and displays the parent of the first pivot item in the pivot table.

**Русский:** Этот скрипт заполняет лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных и отображает родителя первого элемента сводной таблицы.

```javascript
// JavaScript (OnlyOffice) Code

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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
	pages: ['Style'],
	rows: 'Region',
});

// Add data field to the pivot table
pivotTable.AddDataField('Style');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the first pivot field (Style)
var pivotField = pivotTable.GetPivotFields('Style');

// Get the first pivot item
var pivotItem = pivotField.GetPivotItems()[0];

// Set values to display the pivot item's name and its parent
pivotWorksheet.GetRange('A15').SetValue(pivotItem.GetName() + ' parent:');
pivotWorksheet.GetRange('B15').SetValue(pivotItem.GetParent().GetName());
```

```vba
' VBA Code

Sub CreatePivotTableAndDisplayParent()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet

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

    ' Define the data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")

    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"

    ' Create the pivot table
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A3"), TableName:="PivotTable1")

    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlPageField
        .PivotFields("Style").Orientation = xlDataField
        .PivotFields("Style").Function = xlCount
    End With

    ' Get the first pivot field (Style)
    Dim pf As PivotField
    Set pf = pivotTable.PivotFields("Style")

    ' Get the first pivot item
    Dim pi As PivotItem
    Set pi = pf.PivotItems(1)

    ' Display the pivot item's name and its parent
    pivotSheet.Range("A15").Value = pi.Name & " parent:"
    ' VBA does not have a direct method to get the parent of a pivot item
    ' This is a placeholder as VBA PivotItem does not support GetParent
    pivotSheet.Range("B15").Value = "N/A"
End Sub
```