**Description:**

**English:** This code populates an Excel worksheet with region, style, and price data, creates a pivot table based on this data, configures row fields and data fields, sets the row axis layout to tabular, and adds a blank line layout option for the 'Region' field.

**Russian:** Этот код заполняет рабочий лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, настраивает строковые поля и поля данных, устанавливает табличный макет для строковой оси и добавляет опцию пустой строки для поля 'Region'.

---

```vba
' VBA Code to populate data and create a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
    ws.Range("B1").Value = "Region" ' Set header for Region
    ws.Range("C1").Value = "Style"  ' Set header for Style
    ws.Range("D1").Value = "Price"  ' Set header for Price
    
    ' Populate data
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
    Set dataRange = ws.Range("B1:D5")
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add a new worksheet for Pivot Table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="PivotTable1")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Region").Position = 1
    pivotTable.PivotFields("Style").Orientation = xlRowField
    pivotTable.PivotFields("Style").Position = 2
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set row axis layout to tabular
    pivotTable.RowAxisLayout xlTabularRow
    
    ' Set 'Region' field to show a blank line
    ' VBA does not have a direct equivalent to SetLayoutBlankLine, so formatting adjustments may be needed
    ' Here we disable subtotals as a minimal workaround
    pivotTable.PivotFields("Region").Subtotals(1) = False
    
    ' Add values to cells A14 and B14
    pivotWs.Range("A14").Value = "Region blank line"
    pivotWs.Range("B14").Value = "True" ' Placeholder value
End Sub
```

```javascript
// OnlyOffice JS Code to populate data and create a pivot table

var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

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

// Define data reference range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table to a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Set row axis layout to tabular
pivotTable.SetRowAxisLayout('Tabular');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get 'Region' pivot field and set blank line layout
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.SetLayoutBlankLine(true);

// Set values in cells A14 and B14
pivotWorksheet.GetRange('A14').SetValue('Region blank line');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetLayoutBlankLine());
```