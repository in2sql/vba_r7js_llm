### Description / Описание

This code sets up data in an Excel worksheet, creates a pivot table based on that data, and configures the pivot table fields.

Этот код устанавливает данные в рабочий лист Excel, создает сводную таблицу на основе этих данных и настраивает поля сводной таблицы.

```vba
' VBA Code to set up data and create a pivot table
Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Style"
    oWorksheet.Range("D1").Value = "Price"
    
    ' Populate Region column
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("B4").Value = "East"
    oWorksheet.Range("B5").Value = "West"
    
    ' Populate Style column
    oWorksheet.Range("C2").Value = "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"
    oWorksheet.Range("C4").Value = "Tee"
    oWorksheet.Range("C5").Value = "Tee"
    
    ' Populate Price column
    oWorksheet.Range("D2").Value = 42.5
    oWorksheet.Range("D3").Value = 35.2
    oWorksheet.Range("D4").Value = 12.3
    oWorksheet.Range("D5").Value = 24.8
    
    ' Define the data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)
    
    ' Add Row fields
    With pivotTable.PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pivotTable.PivotFields("Style")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ' Add Data field
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "#,##0.00"
        .Name = "Sum of Price"
    End With
    
    ' Add Region as a data field again if needed
    pivotTable.AddDataField pivotTable.PivotFields("Region"), "Count of Region", xlCount
End Sub
```

```javascript
// OnlyOffice JS Code to set up data and create a pivot table
function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');
    
    // Populate Region column
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('B4').SetValue('East');
    oWorksheet.GetRange('B5').SetValue('West');
    
    // Populate Style column
    oWorksheet.GetRange('C2').SetValue('Fancy');
    oWorksheet.GetRange('C3').SetValue('Fancy');
    oWorksheet.GetRange('C4').SetValue('Tee');
    oWorksheet.GetRange('C5').SetValue('Tee');
    
    // Populate Price column
    oWorksheet.GetRange('D2').SetValue(42.5);
    oWorksheet.GetRange('D3').SetValue(35.2);
    oWorksheet.GetRange('D4').SetValue(12.3);
    oWorksheet.GetRange('D5').SetValue(24.8);
    
    // Define the data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert a new pivot table worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add Row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add Data field
    pivotTable.AddDataField('Price');
    
    // Get the active pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the 'Style' pivot field
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Add 'Region' as a data field related to the 'Style' field
    pivotField.GetParent().AddDataField('Region'); 
}
```