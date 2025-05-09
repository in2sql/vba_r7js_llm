### Description / Описание

**English:** This script populates an Excel worksheet with sample data for regions, styles, and prices. It then creates a pivot table based on this data in a new worksheet, configures the pivot table fields and layout, and updates some pivot field properties while displaying related information on the pivot worksheet.

**Русский:** Этот скрипт заполняет лист Excel примерными данными для регионов, стилей и цен. Затем он создает сводную таблицу на основе этих данных на новом листе, настраивает поля и макет сводной таблицы, обновляет некоторые свойства полей сводной таблицы и выводит сопутствующую информацию на лист с сводной таблицей.

```vba
' VBA code equivalent to the OnlyOffice JS script

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
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create pivot table
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Set row layout to Tabular
        .RowAxisLayout xlTabularRow
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Get the pivot field 'Style'
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Write information on the pivot sheet
    pivotSheet.Range("A12").Value = "Style field value"
    pivotSheet.Range("B12").Value = pivotField.Name
    
    ' Set new name for the pivot field
    pivotSheet.Range("A14").Value = "New Style field name"
    pivotField.Name = "My name"
    pivotSheet.Range("B14").Value = pivotField.Name
    
    ' Get source name of the pivot field
    pivotSheet.Range("A15").Value = "Source Style field name"
    pivotSheet.Range("B15").Value = pivotField.SourceName
End Sub
```

```javascript
// OnlyOffice JS equivalent to the provided script

// Function to create and configure pivot table
function createPivotTable() {
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
    
    // Insert pivot table in a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Set row axis layout to Tabular
    pivotTable.SetRowAxisLayout("Tabular", false);
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get pivot field 'Style'
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Write information on the pivot sheet
    pivotWorksheet.GetRange('A12').SetValue('Style field value');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetName());
    
    // Set new name for the pivot field
    pivotWorksheet.GetRange('A14').SetValue('New Style field name');
    pivotField.SetName('My name');
    pivotWorksheet.GetRange('B14').SetValue(pivotField.GetName());
    
    // Get source name of the pivot field
    pivotWorksheet.GetRange('A15').SetValue('Source Style field name');
    pivotWorksheet.GetRange('B15').SetValue(pivotField.GetSourceName());
}

// Call the function to execute
createPivotTable();
```