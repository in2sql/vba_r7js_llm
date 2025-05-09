## Description / Описание

**English:**  
This script populates an active worksheet with regional sales data, creates a pivot table based on this data, and lists the unique regions from the pivot table.

**Russian:**  
Этот скрипт заполняет активный лист данными о продажах по регионам, создает сводную таблицу на основе этих данных и перечисляет уникальные регионы из сводной таблицы.

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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to the pivot table
pivotTable.AddFields({
	rows: 'Region',
	columns: 'Style',
});

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Add the Price field as data in the pivot table
pivotTable.AddDataField('Price');

// Get the pivot field for Region
var pivotField = pivotTable.GetPivotFields('Region');

// Retrieve all unique pivot items for Region
var pivotItems = pivotField.GetPivotItems();

// Set header for pivot items list
pivotWorksheet.GetRange('A10').SetValue('Region pivot items')

// Loop through pivot items and list them
for (var i = 0; i < pivotItems.length; i += 1) {
	pivotWorksheet.GetRangeByNumber(9 + i, 1).SetValue(pivotItems[i].GetName());
}
```

### Excel VBA Code

```vba
Sub CreateSalesPivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim lastRow As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Populate Region column
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Populate Style column
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Populate Price column
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add Row and Column fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
    End With
    
    ' Add Price as data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get unique Regions from pivot table
    Dim pivotField As PivotField
    Dim pivotItem As PivotItem
    Dim i As Integer
    
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set header for pivot items list
    pivotWs.Range("A10").Value = "Region pivot items"
    
    ' List unique regions starting from A11
    i = 0
    For Each pivotItem In pivotField.PivotItems
        pivotWs.Cells(10 + i + 1, 1).Value = pivotItem.Name
        i = i + 1
    Next pivotItem
End Sub
```