### Description
**English:** This script populates an Excel worksheet with specified data, creates a pivot table based on that data, configures the pivot table fields and subtotals, and then outputs the subtotal information to the worksheet.

**Russian:** Этот скрипт заполняет рабочий лист Excel указанными данными, создает на основе этих данных сводную таблицу, настраивает поля и подитоги сводной таблицы, а затем выводит информацию о подитогах на рабочий лист.

```vba
' VBA Code Equivalent to the OnlyOffice JS Code

Sub CreatePivotTable()
    ' Set the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"

    ' Set data
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
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")

    ' Add a new worksheet for pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"

    ' Create Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="PivotTable1")

    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlColumnField
        .PivotFields("Region").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With

    ' Set subtotals for Region to Count
    With pivotTable.PivotFields("Region")
        .Subtotals(1) = True  ' Count
        Dim i As Integer
        For i = 2 To 12
            .Subtotals(i) = False
        Next i
    End With

    ' Get subtotals and write to worksheet
    Dim subtotals As Collection
    Set subtotals = New Collection
    Dim pi As PivotItem
    For Each pi In pivotTable.PivotFields("Region").PivotItems
        subtotals.Add pi.DataRange.Cells(1).PivotCell.DataField.Function, pi.Name
    Next pi

    pivotWs.Range("A11").Value = "Region subtotals"
    Dim row As Integer
    row = 12
    Dim key As Variant
    For Each key In Array("East", "West")
        pivotWs.Cells(row, 1).Value = key
        pivotWs.Cells(row, 2).Value = Application.WorksheetFunction.CountIf(ws.Range("B2:B5"), key)
        row = row + 1
    Next key
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent to the VBA Code

function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();

    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');

    // Set data
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

    // Insert a new pivot table in a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

    // Add fields to pivot table
    pivotTable.AddFields({
        columns: ['Style'],
        rows: 'Region',
    });

    // Add data field
    pivotTable.AddDataField('Price');

    // Get pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();

    // Get pivot field 'Region'
    var pivotField = pivotTable.GetPivotFields('Region');

    // Set subtotals to Count
    pivotField.SetSubtotals({
        Count: true,
    });

    // Get subtotals
    var subtotals = pivotField.GetSubtotals();

    // Write subtotal information to worksheet
    pivotWorksheet.GetRange('A11').SetValue('Region subtotals');
    let k = 12;
    for (var region in subtotals) {
        pivotWorksheet.GetRangeByNumber(k, 0).SetValue(region);
        pivotWorksheet.GetRangeByNumber(k, 1).SetValue(subtotals[region]);
        k++;
    }
}
```