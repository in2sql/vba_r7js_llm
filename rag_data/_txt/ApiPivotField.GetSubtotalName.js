**Description | Описание**

This script populates a worksheet with region, style, and price data, then creates a pivot table that summarizes prices by region and style, and sets a custom subtotal name for the 'Region' field.  
Этот скрипт заполняет лист данными о регионе, стиле и цене, затем создает сводную таблицу, которая суммирует цены по регионам и стилям, и устанавливает пользовательское название подытога для поля 'Region'.

```vba
' VBA Code to populate worksheet and create pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range

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
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Create pivot table
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="SalesPivotTable")

    ' Add fields to pivot table
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlColumnField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Set custom subtotal name for 'Region' field
    With pt.PivotFields("Region")
        .Subtotals(1) = False ' Disable default subtotals
        .Function = xlSum
        .Name = "Region subtotal name"
    End With

    ' Set values in pivot worksheet
    pivotWs.Range("A14").Value = "Region subtotal name"
    pivotWs.Range("B14").Value = pt.PivotFields("Region").Subtotals(1)
End Sub
```

```javascript
// OnlyOffice JS Code to populate worksheet and create pivot table

function main(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set headers
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
    
    // Define data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert pivot table on a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row and column fields
    pivotTable.AddFields({
        columns: ['Region', 'Style'],
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get the pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the 'Region' pivot field
    var pivotField = pivotTable.GetPivotFields('Region');
    
    // Set custom subtotal name
    pivotField.SetSubtotalName('My name');
    
    // Set values in the pivot worksheet
    pivotWorksheet.GetRange('A14').SetValue('Region subtotal name');
    pivotWorksheet.GetRange('B14').SetValue(pivotField.GetSubtotalName());
}
```