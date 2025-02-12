# Description / Описание

**English:**  
This code sets headers and data in specific cells, inserts three pivot tables based on a data range, and adds the 'Price' field to each pivot table.

**Russian:**  
Этот код устанавливает заголовки и данные в определенные ячейки, вставляет три сводные таблицы на основе диапазона данных и добавляет поле 'Price' в каждую сводную таблицу.

```vba
' VBA code
Sub SetupWorksheetAndPivot()
    ' Get active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set header values
    oWorksheet.Range("B1").Value = "Region"
    oWorksheet.Range("C1").Value = "Price"
    
    ' Set data values
    oWorksheet.Range("B2").Value = "East"
    oWorksheet.Range("B3").Value = "West"
    oWorksheet.Range("C2").Value = 42.5
    oWorksheet.Range("C3").Value = 35.2
    
    ' Define data range
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:C3")
    
    ' Insert three pivot tables on new worksheets
    Dim i As Integer
    For i = 1 To 3
        Dim pCache As PivotCache
        Set pCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)
        
        Dim pSheet As Worksheet
        Set pSheet = ThisWorkbook.Worksheets.Add
        pSheet.Name = "PivotTable" & i
        
        Dim pTable As PivotTable
        Set pTable = pCache.CreatePivotTable(TableDestination:=pSheet.Range("A3"), TableName:="PivotTable" & i)
        
        ' Add data field 'Price'
        pTable.AddDataField pTable.PivotFields("Price"), "Sum of Price", xlSum
    Next i
End Sub
```

```javascript
// OnlyOffice JS code
function setupWorksheetAndPivot() {
    // Get active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set header values
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Price');
    
    // Set data values
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('C2').SetValue(42.5);
    oWorksheet.GetRange('C3').SetValue(35.2);
    
    // Define data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
    
    // Insert three pivot tables on new worksheets
    for (var i = 0; i < 3; i++) {
        Api.InsertPivotNewWorksheet(dataRef);
    }
    
    // Add data field 'Price' to all pivot tables
    Api.GetAllPivotTables().forEach(function (pivot) {
        pivot.AddDataField('Price');
    });
}
```