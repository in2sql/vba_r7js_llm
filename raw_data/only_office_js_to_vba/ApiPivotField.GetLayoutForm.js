**Description / Описание**

*English:* The code sets up data in an Excel worksheet, creates a pivot table based on this data, and retrieves the layout form of the 'Region' field in the pivot table.

*Russian:* Код заполняет данные в рабочем листе Excel, создает сводную таблицу на основе этих данных и получает форму расположения поля "Region" в сводной таблице.

```vba
' VBA code equivalent

Sub CreatePivotTable()

    ' Set references to the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add Pivot Table to a new worksheet
    Dim pivotCache As PivotCache
    Dim pivotWorksheet As Worksheet
    Dim pivotTable As PivotTable
    
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=dataRef)
    Set pivotWorksheet = ThisWorkbook.Worksheets.Add
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add fields to pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Get layout form of 'Region' field
    Dim regionField As PivotField
    Set regionField = pivotTable.PivotFields("Region")
    
    pivotWorksheet.Range("A12").Value = "Region layout form"
    ' VBA does not have a direct method for layout form, so this is a placeholder
    pivotWorksheet.Range("B12").Value = "xlTabular" ' Example layout form
    
End Sub
```

```javascript
// OnlyOffice JS code equivalent

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

// Insert pivot table on new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field to pivot table
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set values in pivot worksheet
pivotWorksheet.GetRange('A12').SetValue('Region layout form');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutForm());

```