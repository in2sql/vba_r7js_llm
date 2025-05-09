**Description:**
- **English:** This script populates a worksheet with regional sales data, creates a pivot table to summarize the data by Region and Style, and then retrieves the name of the "Style" field from the pivot table.
- **Russian:** Этот скрипт заполняет лист данными о продажах по регионам, создает сводную таблицу для суммирования данных по региону и стилю, а затем извлекает название поля "Стиль" из сводной таблицы.

```vba
' VBA Code to populate data, create a pivot table, and retrieve a pivot field name

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
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
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Retrieve the "Style" pivot field name
    Set pivotField = pivotTable.PivotFields("Style")
    pivotWs.Range("A12").Value = "The Style field name"
    pivotWs.Range("B12").Value = pivotField.Name
End Sub
```

```javascript
// JavaScript Code to populate data, create a pivot table, and retrieve a pivot field name using OnlyOffice API

function createPivotTable() {
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
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get the active worksheet (pivot table sheet)
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the "Style" pivot field
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Set values in the pivot worksheet
    pivotWorksheet.GetRange('A12').SetValue('The Style field name');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetName());
}

// Execute the function
createPivotTable();
```