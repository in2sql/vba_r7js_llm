### Description / Описание
This script populates an Excel worksheet with sample data and creates a pivot table based on that data.

Этот скрипт заполняет лист Excel примерными данными и создает сводную таблицу на основе этих данных.

```vba
' VBA Code to populate worksheet and create a pivot table

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

    ' Define the data range
    Set dataRange = ws.Range("B1:D5")

    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"

    ' Create Pivot Cache
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Create Pivot Table
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")

    ' Add Row Fields
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField

        ' Add Data Field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Add another Data Field
    With pt.PivotFields("Region")
        .Orientation = xlDataField
        .Function = xlCount
        .Name = "Count of Region"
    End With
End Sub
```

```javascript
// OnlyOffice JS Code to populate worksheet and create a pivot table

function main(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();

    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');

    // Set data for Region
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('B4').SetValue('East');
    oWorksheet.GetRange('B5').SetValue('West');

    // Set data for Style
    oWorksheet.GetRange('C2').SetValue('Fancy');
    oWorksheet.GetRange('C3').SetValue('Fancy');
    oWorksheet.GetRange('C4').SetValue('Tee');
    oWorksheet.GetRange('C5').SetValue('Tee');

    // Set data for Price
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

    // Add data field to the pivot table
    pivotTable.AddDataField('Price');

    // Get the active worksheet where pivot table is created
    var pivotWorksheet = Api.GetActiveSheet();

    // Get the pivot field for 'Style'
    var pivotField = pivotTable.GetPivotFields('Style');

    // Add 'Region' as a data field in the pivot table
    pivotField.GetParent().AddDataField('Region'); 
}
```