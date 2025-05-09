**Description:**

*English: This code sets specific cell values in a worksheet, creates a pivot table based on a data range, adds 'Region' as a row field and 'Price' as a data field to the pivot table, and refreshes all pivot tables.*

*Russian: Этот код устанавливает значения определенных ячеек на рабочем листе, создает сводную таблицу на основе диапазона данных, добавляет «Region» в качестве строкового поля и «Price» в качестве поля данных в сводную таблицу, а также обновляет все сводные таблицы.*

---

**VBA Code:**
```vba
' This VBA code sets cell values, creates a pivot table, adds fields, and refreshes all pivot tables
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotTable As PivotTable
    Dim dataRange As Range
    Dim pivotName As String
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in specific cells
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Price"
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("C2").Value = 42.5
    ws.Range("C3").Value = 35.2
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:C3")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotName = "PivotTable1"
    
    ' Create the pivot table
    Set pivotTable = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange, TableDestination:=pivotWs.Range("A3"), TableName:=pivotName)
    
    ' Add 'Region' to the row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        ' Add 'Price' to the data fields
        .PivotFields("Price").Orientation = xlDataField
    End With
    
    ' Refresh all pivot tables
    ThisWorkbook.RefreshAll
End Sub
```

---

**JavaScript Code (OnlyOffice API):**
```javascript
// This JS code sets cell values, creates a pivot table, adds fields, and refreshes all pivot tables
function createPivotTable(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set values in specific cells
    oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
    oWorksheet.GetRange('C1').SetValue('Price');  // Set header for Price
    oWorksheet.GetRange('B2').SetValue('East');   // Set Region value East
    oWorksheet.GetRange('B3').SetValue('West');   // Set Region value West
    oWorksheet.GetRange('C2').SetValue(42.5);     // Set Price value 42.5
    oWorksheet.GetRange('C3').SetValue(35.2);     // Set Price value 35.2
    
    // Define the data range for the pivot table
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
    
    // Insert a new pivot table on a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add 'Region' as row field in pivot table
    Api.GetPivotByName(pivotTable.GetName()).AddFields({
        rows: 'Region',
    });
    
    // Add 'Price' as data field in pivot table
    Api.GetPivotByName(pivotTable.GetName()).AddDataField('Price');
    
    // Refresh all pivot tables
    Api.RefreshAllPivots();
}
```