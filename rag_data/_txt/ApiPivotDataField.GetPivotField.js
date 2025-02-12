### Description / Описание

**English:**  
This script sets specific values in cells B1 to D5, creates a pivot table based on this data, adds row fields for 'Region' and 'Style', adds 'Price' as data fields, and retrieves the indices of the 'Sum of Price' data field and its corresponding pivot field.

**Russian:**  
Этот скрипт устанавливает определенные значения в ячейки B1 до D5, создает сводную таблицу на основе этих данных, добавляет строковые поля для 'Region' и 'Style', добавляет 'Price' как поля данных и получает индексы поля данных 'Sum of Price' и соответствующего ему сводного поля.

```vba
' VBA Code Equivalent

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataField As PivotField
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Set values in cells B1 to D1
    ws.Range("B1").Value = "Region" ' Set header for Region
    ws.Range("C1").Value = "Style"  ' Set header for Style
    ws.Range("D1").Value = "Price"  ' Set header for Price
    
    ' Set values in column B
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set values in column C
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set values in column D
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add data fields twice for 'Price'
    With pivotTable
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
        .AddDataField .PivotFields("Price"), "Sum of Price 2", xlSum
    End With
    
    ' Retrieve data field index
    Set dataField = pivotTable.PivotFields("Sum of Price")
    
    ' Set values in the pivot worksheet
    pivotWs.Range("A15").Value = "Sum of Price position:"
    pivotWs.Range("B15").Value = dataField.Position
    
    pivotWs.Range("A16").Value = "Price position:"
    pivotWs.Range("B16").Value = dataField.DataRange.Column
End Sub
```

```javascript
// OnlyOffice JS Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data fields twice for 'Price'
pivotTable.AddDataField('Price');
pivotTable.AddDataField('Price');

// Get the active sheet where pivot table is inserted
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field 'Sum of Price'
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A15').SetValue('Sum of Price position:');
pivotWorksheet.GetRange('B15').SetValue(dataField.GetIndex());

pivotWorksheet.GetRange('A16').SetValue('Price position:');
pivotWorksheet.GetRange('B16').SetValue(dataField.GetPivotField().GetIndex());
```