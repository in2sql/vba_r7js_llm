```plaintext
// Description (English):
// This script sets up a worksheet with region and price data, creates three pivot tables based on the data range, 
// and adds the 'Price' field to each pivot table as a data field.

// Описание (Russian):
// Этот скрипт настраивает рабочий лист с данными о регионе и цене, создает три сводные таблицы на основе диапазона данных 
// и добавляет поле 'Price' в каждую сводную таблицу в качестве поля данных.
```

```vba
' VBA Code
Sub CreatePivotTables()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in cells
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Price"
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("C2").Value = 42.5
    ws.Range("C3").Value = 35.2
    
    ' Define the data range
    Set dataRange = ws.Range("B1:C3")
    
    ' Create a PivotCache from the data range
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Insert three pivot tables
    For i = 1 To 3
        ' Add a new worksheet for each pivot table
        Dim pivotSheet As Worksheet
        Set pivotSheet = ThisWorkbook.Worksheets.Add
        pivotSheet.Name = "PivotTable" & i
        
        ' Create the PivotTable
        Set pivotTable = pivotCache.CreatePivotTable( _
            TableDestination:=pivotSheet.Range("A3"), _
            TableName:="PivotTable" & i)
        
        ' Add the 'Price' field as a data field
        With pivotTable
            .PivotFields("Price").Orientation = xlDataField
            .PivotFields("Price").Function = xlSum
            .PivotFields("Price").Name = "Sum of Price"
        End With
    Next i
End Sub
```

```javascript
// OnlyOffice JS Code
// This script sets up a worksheet with region and price data, creates three pivot tables based on the data range,
// and adds the 'Price' field to each pivot table as a data field.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet

// Set values in cells
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Price');  // Set header for Price
oWorksheet.GetRange('B2').SetValue('East');   // Set value East in B2
oWorksheet.GetRange('B3').SetValue('West');   // Set value West in B3
oWorksheet.GetRange('C2').SetValue(42.5);     // Set value 42.5 in C2
oWorksheet.GetRange('C3').SetValue(35.2);     // Set value 35.2 in C3

// Define the data range for pivot tables
var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");

// Insert three new pivot tables on separate worksheets
for (var i = 0; i < 3; i++) {
    Api.InsertPivotNewWorksheet(dataRef); // Insert a new pivot table worksheet
}

// Get all pivot tables and add 'Price' as a data field
Api.GetAllPivotTables().forEach(function (pivot) {
    pivot.AddDataField('Price'); // Add the 'Price' field to the pivot table
});
```