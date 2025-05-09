**Description:**
This script populates an active worksheet with data, creates a pivot table based on that data, and sets specific values from the pivot table into designated cells.
Этот скрипт заполняет активный лист данными, создает сводную таблицу на основе этих данных и устанавливает определенные значения из сводной таблицы в указанные ячейки.

```javascript
// JavaScript Code

// Get the active sheet
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

// Define data range
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active sheet containing the pivot table
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Style' field from the pivot table
var pivotField = pivotTable.GetPivotFields('Style');

// Set values based on the pivot table
pivotWorksheet.GetRange('A12').SetValue('The Style field value');
pivotWorksheet.GetRange('B12').SetValue(pivotField.GetValue());
```

```vba
' VBA Code

' Description:
' This macro populates an active worksheet with data, creates a pivot table based on that data,
' and sets specific values from the pivot table into designated cells.
'
' Описание:
' Этот макрос заполняет активный лист данными, создает сводную таблицу на основе этих данных
' и устанавливает определенные значения из сводной таблицы в указанные ячейки.

Sub CreatePivotTable()
    Dim oWorksheet As Worksheet
    ' Get the active sheet
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
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = oWorksheet.Range("B1:D5")
    
    ' Create a new worksheet for the pivot table
    Dim pivotWorksheet As Worksheet
    Set pivotWorksheet = Worksheets.Add
    
    ' Create the Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotWorksheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Add row fields to the pivot table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
    End With
    
    ' Get the 'Style' field value
    Dim styleValue As String
    styleValue = pivotTable.PivotFields("Style").CurrentPage
    
    ' Set values based on the pivot table
    pivotWorksheet.Range("A12").Value = "The Style field value"
    pivotWorksheet.Range("B12").Value = styleValue
End Sub
```