### Code Description / Описание кода

This script populates an Excel worksheet with region, style, and price data, creates a pivot table based on this data, configures the pivot table fields, and updates the pivot worksheet with specific values.

Этот скрипт заполняет лист Excel данными о регионе, стиле и цене, создает сводную таблицу на основе этих данных, настраивает поля сводной таблицы и обновляет сводный лист с определенными значениями.

```vba
' VBA Code
Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set header values
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data values
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Tee"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Create a new worksheet for Pivot Table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
    
    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add data field
    With pivotTable
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Set ShowAllItems for 'Style' field
    pivotTable.PivotFields("Style").ShowAllItems = True
    
    ' Set values on pivot worksheet
    pivotWorksheet.Range("A12").Value = "Style get show all items"
    pivotWorksheet.Range("B12").Value = pivotTable.PivotFields("Style").ShowAllItems
End Sub
```

```javascript
// JS Code
function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set header values
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');
    
    // Set data values
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('B4').SetValue('East');
    oWorksheet.GetRange('B5').SetValue('West');
    
    oWorksheet.GetRange('C2').SetValue('Fancy');
    oWorksheet.GetRange('C3').SetValue('Tee');
    oWorksheet.GetRange('C4').SetValue('Tee');
    oWorksheet.GetRange('C5').SetValue('Tee');
    
    oWorksheet.GetRange('D2').SetValue(42.5);
    oWorksheet.GetRange('D3').SetValue(35.2);
    oWorksheet.GetRange('D4').SetValue(12.3);
    oWorksheet.GetRange('D5').SetValue(24.8);
    
    // Define data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert Pivot Table in a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Set ShowAllItems for 'Style' field
    pivotField.SetShowAllItems(true);
    
    // Set values on pivot worksheet
    pivotWorksheet.GetRange('A12').SetValue('Style get show all items');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetShowAllItems());
}
```