**Description / Описание**

This code initializes an Excel worksheet by setting up headers and populating data in specific cells. It then creates a pivot table based on the data range, adds fields to the pivot table, and populates a new worksheet with the pivot table's style item values.

Этот код инициализирует рабочий лист Excel, устанавливая заголовки и заполняя данные в определенные ячейки. Затем он создает сводную таблицу на основе диапазона данных, добавляет поля в сводную таблицу и заполняет новый рабочий лист значениями элементов стиля сводной таблицы.

---

```javascript
// OnlyOffice JS Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate 'Region' column
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate 'Style' column
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate 'Price' column
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
	columns: ['Style'],
	rows: 'Region',
});

// Add 'Style' as a data field
pivotTable.AddDataField('Style');

// Get the active worksheet where the pivot table is located
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' field from the pivot table
var pivotField = pivotTable.GetPivotFields('Style');

// Retrieve the pivot items (unique styles)
var pivotItems = pivotField.GetPivotItems();

// Set a header for the style item values
pivotWorksheet.GetRangeByNumber(15, 0).SetValue('Style item values');

// Populate the style item values starting from row 16
for (var i = 0; i < pivotItems.length; i += 1) {
    pivotWorksheet.GetRangeByNumber(15 + i, 1).SetValue(pivotItems[i].GetValue());
}
```

---

```vba
' Excel VBA Code

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Populate 'Region' column
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Populate 'Style' column
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Populate 'Price' column
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add 'Style' to columns
    pivotTable.PivotFields("Style").Orientation = xlColumnField
    
    ' Add 'Region' to rows
    pivotTable.PivotFields("Region").Orientation = xlRowField
    
    ' Add 'Style' to values
    pivotTable.AddDataField pivotTable.PivotFields("Style"), "Count of Style", xlCount
    
    ' Get unique 'Style' items
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Style")
    
    Dim pivotItem As PivotItem
    Dim i As Integer
    i = 16 ' Row 16 (15 + 1)
    
    ' Set header for style item values
    pivotWorksheet.Cells(15, 1).Value = "Style item values"
    
    ' Loop through each pivot item and set the value
    For Each pivotItem In pivotField.PivotItems
        pivotWorksheet.Cells(i, 2).Value = pivotItem.Value
        i = i + 1
    Next pivotItem
End Sub
```