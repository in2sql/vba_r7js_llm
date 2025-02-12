**Description / Описание**

*This script sets up an Excel worksheet with specified headers and data, creates a pivot table based on the data, and retrieves the current page of a pivot field.*

*Этот скрипт настраивает лист Excel с заданными заголовками и данными, создает сводную таблицу на основе данных и получает текущую страницу поля сводной таблицы.*

---

```vba
' VBA Code to replicate OnlyOffice JS functionality

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
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
    
    ' Define the data range for the pivot table
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Sheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTables.Add(PivotCache:=ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef), TableDestination:=pivotSheet.Range("A1"), TableName:="SalesPivot")
    
    ' Add fields to the pivot table
    With pivotTable
        .PivotFields("Style").Orientation = xlPageField
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlDataField
    End With
    
    ' Get the pivot field and set values in the pivot sheet
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Set values in the pivot worksheet
    pivotSheet.Range("A13").Value = "Current Page"
    pivotSheet.Range("B13").Value = pivotField.CurrentPage
End Sub
```

```javascript
// OnlyOffice JS Code to set up worksheet and create a pivot table

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

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new worksheet with the pivot table
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    pages: ['Style'],
    rows: 'Region',
});

// Add a data field
pivotTable.AddDataField('Style');

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set values in the pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Current Page');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetCurrentPage());
```