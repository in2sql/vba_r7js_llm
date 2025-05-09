#### Description / Описание

This code populates specific cells in a worksheet with 'Region' and 'Price' data, sets up a data range, and inserts a pivot table at a designated location.

Данный код заполняет определенные ячейки листа данными 'Region' и 'Price', задает диапазон данных и вставляет сводную таблицу в указанное место.

```vba
' VBA code to replicate OnlyOffice functionality

Sub CreatePivotTable()
    Dim oWorksheet As Worksheet
    Dim dataRange As Range
    Dim pivotRef As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set values in specific cells
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Price"
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("C2").Value = 42.5
    oWorksheet.Range("C3").Value = 35.2

    ' Define data range
    Set dataRange = oWorksheet.Range("B1:C3")

    ' Define pivot table destination
    Set pivotRef = oWorksheet.Range("A7")

    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotRef, TableName:="PivotTable1")
End Sub
```

```javascript
// JavaScript code for OnlyOffice to create pivot table

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in specific cells
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Price');
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('C2').SetValue(42.5);
oWorksheet.GetRange('C3').SetValue(35.2);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");

// Define pivot table destination
var pivotRef = oWorksheet.GetRange('A7');

// Insert pivot table
var pivotTable = Api.InsertPivotExistingWorksheet(dataRef, pivotRef);
```