**Description / Описание**

**English:**  
This script sets up data in an active worksheet, creates a pivot table on a new worksheet, configures the pivot table with specific rows and data fields, and modifies the caption of a pivot table field.

**Russian:**  
Этот скрипт заполняет данные в активном листе, создает сводную таблицу на новом листе, настраивает сводную таблицу с определенными строками и полями данных, а также изменяет заголовок поля сводной таблицы.

---

**VBA Code**

```vba
' VBA code to set up data and create a pivot table

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotTable As PivotTable
    Dim pivotCache As PivotCache
    Dim lastRow As Long, lastCol As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate headers
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
    
    ' Define the data range
    Set dataRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add Row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .RowAxisLayout xlTabularRow
        
        ' Add Data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Modify caption of 'Style' field
    With pivotTable.PivotFields("Style")
        pivotWs.Range("A12").Value = "Style field caption"
        pivotWs.Range("B12").Value = .Caption
        
        pivotWs.Range("A14").Value = "New Style field caption"
        .Caption = "My caption"
        pivotWs.Range("B14").Value = .Caption
    End With
End Sub
```

---

**OnlyOffice JavaScript Code**

```javascript
// JavaScript code to set up data and create a pivot table using OnlyOffice API

function main() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set header values
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Style');
    oWorksheet.GetRange('D1').SetValue('Price');
    
    // Set Region data
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('B4').SetValue('East');
    oWorksheet.GetRange('B5').SetValue('West');
    
    // Set Style data
    oWorksheet.GetRange('C2').SetValue('Fancy');
    oWorksheet.GetRange('C3').SetValue('Fancy');
    oWorksheet.GetRange('C4').SetValue('Tee');
    oWorksheet.GetRange('C5').SetValue('Tee');
    
    // Set Price data
    oWorksheet.GetRange('D2').SetValue(42.5);
    oWorksheet.GetRange('D3').SetValue(35.2);
    oWorksheet.GetRange('D4').SetValue(12.3);
    oWorksheet.GetRange('D5').SetValue(24.8);
    
    // Define the data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert a new pivot table on a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Set row axis layout to Tabular
    pivotTable.SetRowAxisLayout("Tabular", false);
    
    // Add Price as data field
    pivotTable.AddDataField('Price');
    
    // Get the pivot worksheet
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the 'Style' pivot field
    var pivotField = pivotTable.GetPivotFields('Style');
    
    // Set and display original caption
    pivotWorksheet.GetRange('A12').SetValue('Style field caption');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetCaption());
    
    // Change and display new caption
    pivotWorksheet.GetRange('A14').SetValue('New Style field caption');
    pivotField.SetCaption('My caption');
    pivotWorksheet.GetRange('B14').SetValue(pivotField.GetCaption());
}
```