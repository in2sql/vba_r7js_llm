# Code Description / Описание кода

**English:**

This code populates an Excel worksheet with region, style, and price data, creates a pivot table on a new worksheet, adds the 'Price' as a data field, sets 'Region' as column fields and 'Style' as row fields, and lists the pivot column fields starting at cell A9.

**Russian:**

Этот код заполняет рабочий лист Excel данными о регионе, стиле и цене, создает сводную таблицу на новом листе, добавляет «Цена» как поле данных, устанавливает «Регион» как поля колонок и «Стиль» как поля строк, а также перечисляет поля колонок сводной таблицы, начиная с ячейки A9.

```vba
' VBA Code: Populate worksheet, create pivot table, and list column fields

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create pivot table on a new sheet
    Dim pvtCache As PivotCache
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    Dim pvtSheet As Worksheet
    Set pvtSheet = ThisWorkbook.Worksheets.Add
    Dim pvt As PivotTable
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=pvtSheet.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add 'Price' as data field
    pvt.AddDataField pvt.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set 'Region' as column field and 'Style' as row field
    pvt.PivotFields("Region").Orientation = xlColumnField
    pvt.PivotFields("Style").Orientation = xlRowField
    
    ' List column fields starting at A9
    pvtSheet.Range("A9").Value = "Column Fields"
    Dim fld As PivotField
    Dim i As Integer
    i = 0
    For Each fld In pvt.ColumnFields
        pvtSheet.Cells(9 + i, 1).Value = fld.Name
        i = i + 1
    Next fld
End Sub
```

```javascript
// OnlyOffice JS Code: Populate worksheet, create pivot table, and list column fields

// Get active sheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set data
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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Create pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Set 'Region' as column field and 'Style' as row field
pivotTable.AddFields({
    columns: 'Region',
    rows: 'Style',
});

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();
pivotWorksheet.GetRange('A9').SetValue('Column Fields');

// List column fields
var pivotFields = pivotTable.GetColumnFields();
for (var i = 0; i < pivotFields.length; i += 1) {
    var cell = pivotWorksheet.GetRangeByNumber(8 + i, 1);
    cell.SetValue(pivotFields[i].GetName());
}
```