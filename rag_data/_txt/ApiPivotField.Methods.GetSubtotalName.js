**Description / Описание**

English: This code initializes a worksheet by setting headers and data for "Region," "Style," and "Price." It then creates a pivot table from the specified data range, adds "Region" and "Style" as column fields, and "Price" as a data field. Finally, it modifies the subtotal name for the "Region" field and displays it in specific cells.

Russian: Этот код инициализирует рабочий лист, устанавливая заголовки и данные для "Region" (Регион), "Style" (Стиль) и "Price" (Цена). Затем он создает сводную таблицу из указанного диапазона данных, добавляет "Region" и "Style" в качестве полей столбцов, а "Price" в качестве поля данных. В конце он изменяет имя подитога для поля "Region" и отображает его в определенных ячейках.

---

```vba
' VBA Code to replicate the OnlyOffice JS functionality

Sub CreateWorksheetAndPivot()
    Dim ws As Worksheet
    Dim pivotWS As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set Region data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ' Set Style data
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ' Set Price data
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range for the pivot table
    Set dataRange = ws.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Set pivotWS = ThisWorkbook.Worksheets.Add
    pivotWS.Name = "PivotTableSheet"
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWS.Range("A3"), _
        TableName:="PivotTable1")
    
    ' Add Region and Style as column fields
    pivotTable.PivotFields("Region").Orientation = xlColumnField
    pivotTable.PivotFields("Style").Orientation = xlColumnField
    
    ' Add Price as data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Modify the subtotal name for the Region field
    Set pivotField = pivotTable.PivotFields("Region")
    pivotField.Subtotals(1) = False ' Disable default subtotals
    pivotField.Function = xlSum
    pivotField.Name = "My name"
    
    ' Display the subtotal name in specific cells
    pivotWS.Range("A14").Value = "Region subtotal name"
    pivotWS.Range("B14").Value = pivotField.Name
End Sub
```

```javascript
// OnlyOffice JS Code equivalent

function createWorksheetAndPivot(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set headers
    oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
    oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
    oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price
    
    // Set Region data
    oWorksheet.GetRange('B2').SetValue('East');    // Region East
    oWorksheet.GetRange('B3').SetValue('West');    // Region West
    oWorksheet.GetRange('B4').SetValue('East');    // Region East
    oWorksheet.GetRange('B5').SetValue('West');    // Region West
    
    // Set Style data
    oWorksheet.GetRange('C2').SetValue('Fancy');   // Style Fancy
    oWorksheet.GetRange('C3').SetValue('Fancy');   // Style Fancy
    oWorksheet.GetRange('C4').SetValue('Tee');     // Style Tee
    oWorksheet.GetRange('C5').SetValue('Tee');     // Style Tee
    
    // Set Price data
    oWorksheet.GetRange('D2').SetValue(42.5);      // Price 42.5
    oWorksheet.GetRange('D3').SetValue(35.2);      // Price 35.2
    oWorksheet.GetRange('D4').SetValue(12.3);      // Price 12.3
    oWorksheet.GetRange('D5').SetValue(24.8);      // Price 24.8
    
    // Get the data range for the pivot table
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert a new pivot table on a new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add Region and Style as column fields
    pivotTable.AddFields({
        columns: ['Region', 'Style'],
    });
    
    // Add Price as data field
    pivotTable.AddDataField('Price');
    
    // Get the active sheet where the pivot table is located
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get the Region pivot field
    var pivotField = pivotTable.GetPivotFields('Region');
    
    // Set a custom subtotal name for the Region field
    pivotField.SetSubtotalName('My name');
    
    // Display the subtotal name in specific cells
    pivotWorksheet.GetRange('A14').SetValue('Region subtotal name');
    pivotWorksheet.GetRange('B14').SetValue(pivotField.GetSubtotalName());
}
```