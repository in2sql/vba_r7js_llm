# Description / Описание

**English:**  
This script populates an active worksheet with region, style, and price data, creates a pivot table based on this data, adds row and column fields to the pivot table, and then lists the row field names in the pivot table worksheet.

**Russian:**  
Этот скрипт заполняет активный лист данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, добавляет поля строк и столбцов в сводную таблицу, а затем перечисляет имена полей строк на листе сводной таблицы.

## VBA Code

```vba
Sub CreatePivotTable()
    ' Populate headers
    With ActiveSheet
        .Range("B1").Value = "Region"
        .Range("C1").Value = "Style"
        .Range("D1").Value = "Price"
        
        ' Populate data
        .Range("B2").Value = "East"
        .Range("B3").Value = "West"
        .Range("B4").Value = "East"
        .Range("B5").Value = "West"
        
        .Range("C2").Value = "Fancy"
        .Range("C3").Value = "Fancy"
        .Range("C4").Value = "Tee"
        .Range("C5").Value = "Tee"
        
        .Range("D2").Value = 42.5
        .Range("D3").Value = 35.2
        .Range("D4").Value = 12.3
        .Range("D5").Value = 24.8
    End With
    
    ' Define the data range
    Dim dataRange As Range
    Set dataRange = ActiveSheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the PivotTable
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
        
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A1"), _
        TableName:="SalesPivotTable")
        
    ' Add row and column fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Add description for row fields
    pivotSheet.Range("A9").Value = "Row Fields"
    
    ' List row field names
    Dim rowFields As PivotFields
    Set rowFields = pivotTable.RowFields
    Dim i As Integer
    For i = 1 To rowFields.Count
        pivotSheet.Cells(8 + i, 1).Value = rowFields(i).Name
    Next i
End Sub
```

## OnlyOffice JS Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields
pivotTable.AddFields({
    rows: 'Region',
    columns: 'Style',
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Add description for row fields
pivotWorksheet.GetRange('A9').SetValue('Row Fields');

// Get row fields from the pivot table
var rowFields = pivotTable.GetRowFields();

// List row field names starting from A10
for (var i = 0; i < rowFields.length; i += 1) {
    var cell = pivotWorksheet.GetRangeByNumber(9 + i, 1); // Rows are 0-indexed
    cell.SetValue(rowFields[i].GetName());
}
```