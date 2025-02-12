**Description / Описание**

This code populates data into an active worksheet, creates a pivot table based on the data, adds row and data fields to the pivot table, changes the layout form of the 'Region' field, and outputs the layout form setting.

Этот код заполняет данные в активном листе, создает сводную таблицу на основе этих данных, добавляет строковые и числовые поля в сводную таблицу, изменяет форму отображения поля 'Region' и выводит настройку формы отображения.

```vba
' VBA Code Equivalent

Sub CreatePivotTable()

    ' Populate data in the active worksheet
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
    
    ' Add a new worksheet for Pivot Table
    Dim pivotWS As Worksheet
    Set pivotWS = ThisWorkbook.Worksheets.Add
    pivotWS.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Dim pCache As PivotCache
    Set pCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Dim pTable As PivotTable
    Set pTable = pCache.CreatePivotTable( _
        TableDestination:=pivotWS.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add row fields
    pTable.PivotFields("Region").Orientation = xlRowField
    pTable.PivotFields("Region").Position = 1
    pTable.PivotFields("Style").Orientation = xlRowField
    pTable.PivotFields("Style").Position = 2
    
    ' Add data field
    pTable.AddDataField pTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set layout form for 'Region' field to Tabular
    With pTable.PivotFields("Region")
        .LayoutForm = xlTabular
    End With
    
    ' Output the layout form setting
    pivotWS.Range("A12").Value = "Region layout form"
    pivotWS.Range("B12").Value = "Tabular" ' Since VBA sets it to xlTabular

End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// Get active worksheet
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

// Insert Pivot Table in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set layout form to Tabular
pivotField.SetLayoutForm("Tabular");

// Output the layout form setting
pivotWorksheet.GetRange('A12').SetValue('Region layout form');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutForm()); 
```