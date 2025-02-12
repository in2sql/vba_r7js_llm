**Description**

This script sets up data in an Excel worksheet, creates a pivot table based on that data, and configures the pivot table's fields and layout.

Этот скрипт устанавливает данные в рабочем листе Excel, создает сводную таблицу на основе этих данных и настраивает поля и макет сводной таблицы.

```vba
' Excel VBA Code

Sub CreatePivotTable()
    ' Set header values in B1, C1, D1
    With ActiveSheet
        .Range("B1").Value = "Region"
        .Range("C1").Value = "Style"
        .Range("D1").Value = "Price"

        ' Set data for Region
        .Range("B2").Value = "East"
        .Range("B3").Value = "West"
        .Range("B4").Value = "East"
        .Range("B5").Value = "West"

        ' Set data for Style
        .Range("C2").Value = "Fancy"
        .Range("C3").Value = "Fancy"
        .Range("C4").Value = "Tee"
        .Range("C5").Value = "Tee"

        ' Set data for Price
        .Range("D2").Value = 42.5
        .Range("D3").Value = 35.2
        .Range("D4").Value = 12.3
        .Range("D5").Value = 24.8
    End With

    ' Define the data range
    Dim dataRange As Range
    Set dataRange = ActiveSheet.Range("B1:D5")

    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"

    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTables.Add(PivotCache:=ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange), TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")

    ' Add row fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
    End With

    ' Add data field
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum

    ' Set value in A12
    pivotSheet.Range("A12").Value = "Table Style Row Headers"

    ' Retrieve table style row headers property (example placeholder)
    ' Excel VBA does not have a direct equivalent; setting as True for demonstration
    pivotSheet.Range("B12").Value = True
End Sub
```

```javascript
// OnlyOffice JS Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
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

// Insert a pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields
pivotTable.AddFields({
    rows: ['Region', 'Style']
});

// Add data field
pivotTable.AddDataField('Price');

// Get the pivot table worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Set values in A12 and B12
pivotWorksheet.GetRange('A12').SetValue('Table Style Row Headers');
pivotWorksheet.GetRange('B12').SetValue(pivotTable.GetTableStyleRowHeaders());
```