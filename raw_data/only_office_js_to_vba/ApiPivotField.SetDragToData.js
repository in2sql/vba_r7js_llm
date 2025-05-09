## Description / Описание

**English:**  
This code creates and manipulates an Excel worksheet by setting specific cell values, creating a pivot table in a new worksheet, adding row and column fields, adding a data field, modifying pivot field properties, and updating certain cells based on the pivot table's configuration.

**Russian:**  
Этот код создает и управляет листом Excel, устанавливая значения в определенные ячейки, создавая сводную таблицу на новом листе, добавляя поля строк и столбцов, добавляя поле данных, изменяя свойства полей сводной таблицы и обновляя определенные ячейки на основе конфигурации сводной таблицы.

```vba
' VBA Code Equivalent for OnlyOffice API Example
' VBA Код, эквивалентный примеру OnlyOffice API

Sub CreateAndManipulatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Set data for Region
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Set data for Style
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Set data for Price
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = Worksheets.Add
    pivotWorksheet.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWorksheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add fields to the pivot table
    With pivotTable
        ' Add Style as row field
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Style").Position = 1
        
        ' Add Region as column field
        .PivotFields("Region").Orientation = xlColumnField
        .PivotFields("Region").Position = 1
        
        ' Add Price as data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Attempt to modify pivot field properties (Note: Excel VBA has limited properties compared to OnlyOffice)
    ' Excel VBA does not have a direct equivalent for SetDragToData
    ' This part is illustrative and may require custom implementation
    ' For example, you might hide the field or protect the sheet
    pivotTable.PivotFields("Region").EnableMultiplePageItems = False
    
    ' Update cells based on pivot table configuration
    pivotWorksheet.Range("A13").Value = "Drag to data"
    pivotWorksheet.Range("B13").Value = False ' Assuming SetDragToData(false)
    pivotWorksheet.Range("A14").Value = "Try drag Region to data!"
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent with Comments

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

// Define the data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to the pivot table
pivotTable.AddFields({
    rows: ['Style'],    // Add 'Style' as row field
    columns: 'Region',  // Add 'Region' as column field
});

// Add 'Price' as data field
pivotTable.AddDataField('Price');

// Modify pivot field properties
var pivotWorksheet = Api.GetActiveSheet();
var pivotField = pivotTable.GetPivotFields('Region');

// Disable dragging 'Region' to data area
pivotField.SetDragToData(false);

// Update cells based on pivot table configuration
pivotWorksheet.GetRange('A13').SetValue('Drag to data');
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToData());
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to data!');
```