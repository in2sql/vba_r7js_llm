---

**Description:**

This script populates an Excel sheet with region, style, and price data, creates a pivot table analyzing the data, and sets subtotal configurations for the 'Region' field.

Этот скрипт заполняет лист Excel данными о регионе, стиле и цене, создаёт сводную таблицу для анализа данных и устанавливает настройки подитогов для поля 'Region'.

---

```vba
' VBA Code Equivalent

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A3"), TableName:="SalesPivotTable")
    
    ' Add row fields
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Configure subtotal for 'Region'
    With pivotTable.PivotFields("Region")
        .Subtotals(1) = True ' Enable subtotals
        .LayoutSubtotalLocation = xlAtBottom
        .Subtotals(1) = True
    End With
    
    ' Rename subtotal
    pivotTable.PivotFields("Region").Function = xlSum
    pivotTable.PivotFields("Region").Name = "My name"
    
    ' Set subtotal name in cells
    pivotWorksheet.Range("A14").Value = "Region subtotal name"
    pivotWorksheet.Range("B14").Value = pivotTable.PivotFields("Region").Name
End Sub
```

```javascript
// OnlyOffice JS Code

// Get the active worksheet
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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Configure subtotal for 'Region'
var pivotField = pivotTable.GetPivotFields('Region');
pivotField.SetLayoutSubtotalLocation('Bottom');
pivotField.SetSubtotalName('My name');

// Set subtotal name in cells
pivotWorksheet.GetRange('A14').SetValue('Region subtotal name');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetSubtotalName());
```