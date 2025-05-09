---

**Description / Описание:**

This script sets up a worksheet with data and creates a pivot table based on that data.
Этот скрипт настраивает лист с данными и создает сводную таблицу на их основе.

```javascript
// OnlyOffice JavaScript code

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
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Get the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A14').SetValue('Region layout subtotal location');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetLayoutSubtotalLocation());
```

```vba
' Excel VBA code

' This macro sets up data in the worksheet and creates a pivot table based on that data.

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data values
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
        SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), TableName:="PivotTable1")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With
    
    ' Add data field
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Price"
    End With
    
    ' Get the 'Region' pivot field
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set values in the pivot worksheet
    pivotWs.Range("A14").Value = "Region layout subtotal location"
    
    ' VBA does not have a direct equivalent of GetLayoutSubtotalLocation
    ' This is a placeholder to demonstrate where such a value would be set or retrieved
    pivotWs.Range("B14").Value = "Subtotal Location Not Available"
    
End Sub
```