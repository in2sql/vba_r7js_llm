---

**Description / Описание**

**English:**  
The code sets values in specific cells of the active worksheet, creates a pivot table based on a data range, adds fields to the pivot table, adds and configures data fields, and sets values indicating the orientation of the data field in another worksheet.

**Russian:**  
Код устанавливает значения в определенные ячейки активного листа, создает сводную таблицу на основе диапазона данных, добавляет поля в сводную таблицу, добавляет и настраивает поля данных, а также устанавливает значения, указывающие ориентацию поля данных на другом листе.

---

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers in cells B1, C1, D1
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Set values in column B
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set values in column C
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set values in column D
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add a data field for 'Price'
pivotTable.AddDataField('Price');
var dataField = pivotTable.AddDataField('Price');

// Set the position of the data field
dataField.SetPosition(1);

// Get the active worksheet for the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Set values indicating the orientation of the data field
pivotWorksheet.GetRange('A15').SetValue('Sum of Price2 orientation:');
pivotWorksheet.GetRange('B15').SetValue(dataField.GetOrientation());
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers in cells B1, C1, D1
oWorksheet.Range("B1").Value = "Region"
oWorksheet.Range("C1").Value = "Style"
oWorksheet.Range("D1").Value = "Price"

' Set values in column B
oWorksheet.Range("B2").Value = "East"
oWorksheet.Range("B3").Value = "West"
oWorksheet.Range("B4").Value = "East"
oWorksheet.Range("B5").Value = "West"

' Set values in column C
oWorksheet.Range("C2").Value = "Fancy"
oWorksheet.Range("C3").Value = "Fancy"
oWorksheet.Range("C4").Value = "Tee"
oWorksheet.Range("C5").Value = "Tee"

' Set values in column D
oWorksheet.Range("D2").Value = 42.5
oWorksheet.Range("D3").Value = 35.2
oWorksheet.Range("D4").Value = 12.3
oWorksheet.Range("D5").Value = 24.8

' Define the data range for the pivot table
Dim dataRange As Range
Set dataRange = oWorksheet.Range("B1:D5")

' Insert a new pivot table in a new worksheet
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

Dim pivotWorksheet As Worksheet
Set pivotWorksheet = ThisWorkbook.Worksheets.Add
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")

' Add row fields to the pivot table
With pivotTable
    .PivotFields("Region").Orientation = xlRowField
    .PivotFields("Style").Orientation = xlRowField
End With

' Add a data field for 'Price' and set its position
With pivotTable.PivotFields("Price")
    .Orientation = xlDataField
    .Function = xlSum
    .Position = 1
End With

' Set values indicating the orientation of the data field
pivotWorksheet.Range("A15").Value = "Sum of Price2 orientation:"
pivotWorksheet.Range("B15").Value = pivotTable.PivotFields("Price").Orientation
```