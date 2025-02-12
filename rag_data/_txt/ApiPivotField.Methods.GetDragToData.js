```javascript
// Description: This script populates the active worksheet with data, creates a pivot table from the data, 
// configures the pivot table fields, and sets specific values in the pivot table.
// Описание: Этот скрипт заполняет активный лист данными, создает сводную таблицу из этих данных, 
// настраивает поля сводной таблицы и устанавливает определенные значения в сводной таблице.

// OnlyOffice JavaScript Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Populate Region column
oWorksheet.GetRange('B2').SetValue('East');   // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');   // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');   // Set 'West' in cell B5

// Populate Style column
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set 'Tee' in cell C5

// Populate Price column
oWorksheet.GetRange('D2').SetValue(42.5);     // Set 42.5 in cell D2
oWorksheet.GetRange('D3').SetValue(35.2);     // Set 35.2 in cell D3
oWorksheet.GetRange('D4').SetValue(12.3);     // Set 12.3 in cell D4
oWorksheet.GetRange('D5').SetValue(24.8);     // Set 24.8 in cell D5

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row and column fields to the pivot table
pivotTable.AddFields({
    rows: ['Style'],      // Add 'Style' as row field
    columns: 'Region',    // Add 'Region' as column field
});

// Add 'Price' as a data field in the pivot table
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set values in specific cells of the pivot table worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to data');          // Set 'Drag to data' in cell A13
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToData()); // Set drag to data status in cell B13
```

```vba
' Description: This macro populates the active worksheet with data, creates a pivot table from the data, 
' configures the pivot table fields, and sets specific values in the pivot table.
' Описание: Этот макрос заполняет активный лист данными, создает сводную таблицу из этих данных, 
' настраивает поля сводной таблицы и устанавливает определенные значения в сводной таблице.

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim dataRange As Range
    Dim pivotRange As Range
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"   ' Set 'Region' in cell B1
    ws.Range("C1").Value = "Style"    ' Set 'Style' in cell C1
    ws.Range("D1").Value = "Price"    ' Set 'Price' in cell D1
    
    ' Populate Region column
    ws.Range("B2").Value = "East"      ' Set 'East' in cell B2
    ws.Range("B3").Value = "West"      ' Set 'West' in cell B3
    ws.Range("B4").Value = "East"      ' Set 'East' in cell B4
    ws.Range("B5").Value = "West"      ' Set 'West' in cell B5
    
    ' Populate Style column
    ws.Range("C2").Value = "Fancy"     ' Set 'Fancy' in cell C2
    ws.Range("C3").Value = "Fancy"     ' Set 'Fancy' in cell C3
    ws.Range("C4").Value = "Tee"       ' Set 'Tee' in cell C4
    ws.Range("C5").Value = "Tee"       ' Set 'Tee' in cell C5
    
    ' Populate Price column
    ws.Range("D2").Value = 42.5        ' Set 42.5 in cell D2
    ws.Range("D3").Value = 35.2        ' Set 35.2 in cell D3
    ws.Range("D4").Value = 12.3        ' Set 12.3 in cell D4
    ws.Range("D5").Value = 24.8        ' Set 24.8 in cell D5
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    
    ' Create the pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create the pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add 'Style' to the row fields
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add 'Region' to the column fields
    pivotTable.PivotFields("Region").Orientation = xlColumnField
    
    ' Add 'Price' to the data fields
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Set specific values in the pivot table worksheet
    pivotWs.Range("A13").Value = "Drag to data"   ' Set 'Drag to data' in cell A13
    ' Note: VBA does not have a direct equivalent to GetDragToData, so this part may require additional logic
    ' For demonstration, we'll set a placeholder value
    pivotWs.Range("B13").Value = "DragToDataValue" ' Placeholder for drag to data status
End Sub
```