# Description / Описание
**English:**  
This script populates an Excel worksheet with region, style, and price data, creates a pivot table based on this data, configures the pivot table fields, sets the layout to tabular form, and adds a custom label with the pivot field's blank line setting.

**Russian:**  
Этот скрипт заполняет лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, настраивает поля сводной таблицы, устанавливает табличный формат отображения и добавляет пользовательскую метку с настройкой пустой строки поля сводной таблицы.

```vba
' VBA Code Equivalent
Sub CreatePivotTable()
    ' Declare variables
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
    
    ' Define data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set layout to tabular
    pivotTable.RowAxisLayout xlTabularRow
    
    ' Get the pivot field 'Region'
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Add custom labels
    pivotWs.Range("A14").Value = "Region blank line"
    pivotWs.Range("B14").Value = pivotField.LayoutBlankLine
End Sub
```

```javascript
// JavaScript Code Equivalent (OnlyOffice API)
function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region'); // Set "Region" in B1
    oWorksheet.GetRange('C1').SetValue('Style');  // Set "Style" in C1
    oWorksheet.GetRange('D1').SetValue('Price');  // Set "Price" in D1
    
    // Populate Region data
    oWorksheet.GetRange('B2').SetValue('East');  // Set "East" in B2
    oWorksheet.GetRange('B3').SetValue('West');  // Set "West" in B3
    oWorksheet.GetRange('B4').SetValue('East');  // Set "East" in B4
    oWorksheet.GetRange('B5').SetValue('West');  // Set "West" in B5
    
    // Populate Style data
    oWorksheet.GetRange('C2').SetValue('Fancy'); // Set "Fancy" in C2
    oWorksheet.GetRange('C3').SetValue('Fancy'); // Set "Fancy" in C3
    oWorksheet.GetRange('C4').SetValue('Tee');   // Set "Tee" in C4
    oWorksheet.GetRange('C5').SetValue('Tee');   // Set "Tee" in C5
    
    // Populate Price data
    oWorksheet.GetRange('D2').SetValue(42.5);     // Set 42.5 in D2
    oWorksheet.GetRange('D3').SetValue(35.2);     // Set 35.2 in D3
    oWorksheet.GetRange('D4').SetValue(12.3);     // Set 12.3 in D4
    oWorksheet.GetRange('D5').SetValue(24.8);     // Set 24.8 in D5
    
    // Define data range for pivot table
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert pivot table in a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'], // Add "Region" and "Style" as row fields
    });
    
    // Add data field
    pivotTable.AddDataField('Price'); // Add "Price" as data field
    
    // Set layout to tabular
    pivotTable.SetRowAxisLayout('Tabular'); // Set row layout to tabular
    
    // Get the active sheet (pivot table sheet)
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the pivot field "Region"
    var pivotField = pivotTable.GetPivotFields('Region');
    
    // Add custom labels
    pivotWorksheet.GetRange('A14').SetValue('Region blank line'); // Set label in A14
    pivotWorksheet.GetRange('B14').SetValue(pivotField.GetLayoutBlankLine()); // Set blank line setting in B14
}

// Execute the function
createPivotTable();
```