## Code Description / Описание кода

**English:**  
This code populates a worksheet with sample data, creates a pivot table based on that data, configures the pivot table fields, and updates the name of one of the pivot fields.

**Russian:**  
Этот код заполняет рабочий лист примерными данными, создает сводную таблицу на основе этих данных, настраивает поля сводной таблицы и обновляет имя одного из полей сводной таблицы.

---

### OnlyOffice JS Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange('B1').SetValue('Region');
oWorksheet.GetRange('C1').SetValue('Style');
oWorksheet.GetRange('D1').SetValue('Price');

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');
oWorksheet.GetRange('B3').SetValue('West');
oWorksheet.GetRange('B4').SetValue('East');
oWorksheet.GetRange('B5').SetValue('West');

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');
oWorksheet.GetRange('C3').SetValue('Fancy');
oWorksheet.GetRange('C4').SetValue('Tee');
oWorksheet.GetRange('C5').SetValue('Tee');

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);
oWorksheet.GetRange('D3').SetValue(35.2);
oWorksheet.GetRange('D4').SetValue(12.3);
oWorksheet.GetRange('D5').SetValue(24.8);

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Set the row axis layout to Tabular form
pivotTable.SetRowAxisLayout("Tabular", false);

// Add the Price field as a data field
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' pivot field
var pivotField = pivotTable.GetPivotFields('Style');

// Set and display the original name of the 'Style' field
pivotWorksheet.GetRange('A12').SetValue('Style field name');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetName());

// Update the name of the 'Style' field and display the new name
pivotWorksheet.GetRange('A14').SetValue('New Style field name');
pivotField.SetName('My name');
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetName());
```

---

### Excel VBA Code

```vba
Sub CreatePivotTable()
    ' Declare variables
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim dataRange As Range
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Populate Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Populate Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Populate Price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add Row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        .PivotFields("Style").Orientation = xlRowField
        .PivotFields("Style").Position = 2
    End With
    
    ' Set row layout to Tabular
    pivotTable.RowAxisLayout xlTabularRow
    
    ' Add Data field
    With pivotTable.PivotFields("Price")
        .Orientation = xlDataField
        .Function = xlSum
        .Name = "Sum of Price"
    End With
    
    ' Get the 'Style' pivot field
    Set pivotField = pivotTable.PivotFields("Style")
    
    ' Set and display the original name of the 'Style' field
    pivotWs.Range("A12").Value = "Style field name"
    pivotWs.Range("B12").Value = pivotField.Name
    
    ' Update the name of the 'Style' field and display the new name
    pivotWs.Range("A14").Value = "New Style field name"
    pivotField.Name = "My name"
    pivotWs.Range("B14").Value = pivotField.Name
End Sub
```