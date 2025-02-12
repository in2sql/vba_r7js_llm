**Description / Описание:**
This script sets up a worksheet with specific headers and data, creates a pivot table based on that data, and then populates additional information from the pivot table.  
Этот скрипт настраивает лист с определенными заголовками и данными, создает сводную таблицу на основе этих данных и затем заполняет дополнительную информацию из сводной таблицы.

```javascript
// JavaScript OnlyOffice API code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange('B1').SetValue('Region'); // Set value 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set value 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set value 'Price' in cell D1

// Set data values for Region
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Set data values for Style
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Set data values for Price
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    columns: ['Style'], // Set 'Style' as column field
    rows: 'Region',     // Set 'Region' as row field
});

// Add data field to the pivot table
pivotTable.AddDataField('Style'); // Add 'Style' as data field

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field 'Style'
var pivotField = pivotTable.GetPivotFields('Style');

// Get all items in the 'Style' pivot field
var pivotItems = pivotField.GetPivotItems();

// Set header for style item values
pivotWorksheet.GetRangeByNumber(15, 0).SetValue('Style item values');

// Populate style item values below the header
for (var i = 0; i < pivotItems.length; i += 1) {
    pivotWorksheet.GetRangeByNumber(15 + i, 1).SetValue(pivotItems[i].GetValue());
}
```

```vba
' VBA Equivalent Code

' Description:
' This VBA script sets up a worksheet with specific headers and data, creates a pivot table based on that data,
' and then populates additional information from the pivot table.

Sub CreatePivotTable()
    ' Set references
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = "Region" ' Set value 'Region' in cell B1
    ws.Range("C1").Value = "Style"  ' Set value 'Style' in cell C1
    ws.Range("D1").Value = "Price"  ' Set value 'Price' in cell D1
    
    ' Set data values for Region
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set data values for Style
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set data values for Price
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the range for the pivot table
    Dim dataRange As Range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWs As Worksheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the Pivot Cache
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create the Pivot Table
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="PivotTable1")
    
    ' Add fields to the pivot table
    With pt
        .PivotFields("Style").Orientation = xlColumnField ' Set 'Style' as column field
        .PivotFields("Region").Orientation = xlRowField   ' Set 'Region' as row field
        .AddDataField .PivotFields("Style"), "Count of Style", xlCount ' Add 'Style' as data field
    End With
    
    ' Get the pivot items from 'Style' field
    Dim pf As PivotField
    Set pf = pt.PivotFields("Style")
    
    Dim pi As PivotItem
    Dim i As Integer
    
    ' Set header for style item values
    pivotWs.Cells(15, 1).Value = "Style item values"
    
    ' Populate style item values below the header
    i = 1
    For Each pi In pf.PivotItems
        pivotWs.Cells(15 + i, 2).Value = pi.Name
        i = i + 1
    Next pi
End Sub
```