**Description / Описание**

This script populates an Excel worksheet with data, creates a pivot table based on that data, and displays summary functions related to the pivot table.

Этот скрипт заполняет рабочий лист Excel данными, создает сводную таблицу на основе этих данных и отображает сводные функции, связанные со сводной таблицей.

```vba
' VBA Code to populate data, create a pivot table, and display summary functions

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim sumDataField As PivotField
    Dim countDataField As PivotField
    
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
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
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
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data fields
    Set sumDataField = pivotTable.AddDataField(pivotTable.PivotFields("Price"), "Sum of Price", xlSum)
    Set countDataField = pivotTable.AddDataField(pivotTable.PivotFields("Price"), "Count of Price", xlCount)
    
    ' Display functions
    pivotWs.Range("A15").Value = "Functions:"
    pivotWs.Range("B15").Value = sumDataField.Function
    pivotWs.Range("B16").Value = countDataField.Function
End Sub
```

```javascript
// OnlyOffice JS Code to populate data, create a pivot table, and display summary functions

function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
    oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
    oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price
    
    // Populate Region data
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('B4').SetValue('East');
    oWorksheet.GetRange('B5').SetValue('West');
    
    // Populate Style data
    oWorksheet.GetRange('C2').SetValue('Fancy');
    oWorksheet.GetRange('C3').SetValue('Fancy');
    oWorksheet.GetRange('C4').SetValue('Tee');
    oWorksheet.GetRange('C5').SetValue('Tee');
    
    // Populate Price data
    oWorksheet.GetRange('D2').SetValue(42.5);
    oWorksheet.GetRange('D3').SetValue(35.2);
    oWorksheet.GetRange('D4').SetValue(12.3);
    oWorksheet.GetRange('D5').SetValue(24.8);
    
    // Define data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert a new pivot table in a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields to the pivot table
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add Sum of Price field
    var sumDataField = pivotTable.AddDataField('Price');
    
    // Add Count of Price field and set its function to Count
    var countDataField = pivotTable.AddDataField('Price');
    countDataField.SetFunction('Count');
    
    // Get the pivot table worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Display functions used in the pivot table
    pivotWorksheet.GetRange('A15').SetValue('Functions:');
    pivotWorksheet.GetRange('B15').SetValue(sumDataField.GetFunction());
    pivotWorksheet.GetRange('B16').SetValue(countDataField.GetFunction());
}
```