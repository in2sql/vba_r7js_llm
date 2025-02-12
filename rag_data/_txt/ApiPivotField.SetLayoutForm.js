**Description | Описание**

English: The code populates data into an Excel worksheet, creates a pivot table on a new worksheet using specified fields, sets the layout form of a pivot field, and displays the layout form in designated cells.

Russian: Код заполняет данные в рабочий лист Excel, создает сводную таблицу на новом листе, используя указанные поля, устанавливает форму макета поля сводной таблицы и отображает форму макета в определенных ячейках.

---

```vba
' English: VBA code to populate data, create a pivot table, configure fields, set layout form, and display the layout form.
' Russian: VBA код для заполнения данных, создания сводной таблицы, настройки полей, установки формы макета и отображения формы макета.

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotRange As Range
    Dim pivotTable As PivotTable
    Dim pivotField As PivotField
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set headers
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Style"
    ws.Range("D1").Value = "Price"
    
    ' Set data
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("B4").Value = "East"
    ws.Range("B5").Value = "West"
    
    ws.Range("C2").Value = "Fancy"
    ws.Range("C3").Value = "Fancy"
    ws.Range("C4").Value = "Tee"
    ws.Range("C5").Value = "Tee"
    
    ws.Range("D2").Value = 42.5
    ws.Range("D3").Value = 35.2
    ws.Range("D4").Value = 12.3
    ws.Range("D5").Value = 24.8
    
    ' Define the data range
    Set pivotRange = ws.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Set pivotWs = ThisWorkbook.Worksheets.Add
    pivotWs.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Set pivotTable = pivotWs.PivotTableWizard(SourceType:=xlDatabase, SourceData:=pivotRange)
    
    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Set layout form for the 'Region' field
    Set pivotField = pivotTable.PivotFields("Region")
    pivotField.LayoutForm = xlTabular
    
    ' Display the layout form
    pivotWs.Range("A12").Value = "Region layout form"
    pivotWs.Range("B12").Value = pivotField.LayoutForm
End Sub
```

```javascript
// English: OnlyOffice JavaScript code to populate data, create a pivot table, configure fields, set layout form, and display the layout form.
// Russian: JavaScript код OnlyOffice для заполнения данных, создания сводной таблицы, настройки полей, установки формы макета и отображения формы макета.

function createPivotTable() {
    // Get the active worksheet
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
    
    // Define the data range
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
    
    // Insert pivot table in new worksheet
    var pivotTable = Api.InsertPivotNewWorksheet(dataRef);
    
    // Add row fields
    pivotTable.AddFields({
        rows: ['Region', 'Style'],
    });
    
    // Add data field
    pivotTable.AddDataField('Price');
    
    // Get active sheet for pivot table
    var pivotWorksheet = Api.GetActiveSheet();
    
    // Get pivot field 'Region' and set layout form
    var pivotField = pivotTable.GetPivotFields('Region');
    pivotField.SetLayoutForm("Tabular");
    
    // Display the layout form
    pivotWorksheet.GetRange('A12').SetValue('Region layout form');
    pivotWorksheet.GetRange('B12').SetValue(pivotField.GetLayoutForm());
}
```