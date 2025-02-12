# Code Description / Описание кода

**English:**  
This code populates the active worksheet with region and price data, and creates a pivot table in a new worksheet summarizing the prices by region.

**Russian:**  
Этот код заполняет активный лист данными о регионах и ценах и создает сводную таблицу на новом листе, суммируя цены по регионам.

```javascript
// OnlyOffice JS Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Price');

// Set data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('C2').SetValue(42.5);
oWorksheet.GetRange('C3').SetValue(35.2);

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Region' as a row field
Api.GetPivotByName(pivotTable.GetName()).AddFields({
    rows: 'Region',
});

// Add 'Price' as a data field
Api.GetPivotByName(pivotTable.GetName()).AddDataField('Price');
```

```vba
' Excel VBA Code

Sub CreatePivotTable()
    ' Reference to the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Price"
    
    ' Set data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("C2").Value = 42.5
    ws.Range("C3").Value = 35.2
    
    ' Define data range
    Dim dataRef As Range
    Set dataRef = ws.Range("B1:C3")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWS As Worksheet
    Set pivotWS = ThisWorkbook.Worksheets.Add
    pivotWS.Name = "PivotTableSheet"
    
    ' Create the pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRef)
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWS.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add 'Region' as a row field
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        ' Add 'Price' as a data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
End Sub
```