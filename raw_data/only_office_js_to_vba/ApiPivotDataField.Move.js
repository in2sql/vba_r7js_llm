### Description / Описание

**English:**  
The code sets values in specific cells, creates a pivot table based on this data, adds fields to the pivot table, and moves a data field to rows after 5 seconds.

**Russian:**  
Код устанавливает значения в определенные ячейки, создает сводную таблицу на основе этих данных, добавляет поля в сводную таблицу и перемещает поле данных в строки через 5 секунд.

```vba
' VBA Code Equivalent

Sub CreatePivotTable()
    ' Set values in cells
    With ThisWorkbook.ActiveSheet
        .Range("B1").Value = "Region" ' Set header for Region
        .Range("C1").Value = "Style"  ' Set header for Style
        .Range("D1").Value = "Price"  ' Set header for Price
        
        .Range("B2").Value = "East"
        .Range("B3").Value = "West"
        .Range("B4").Value = "East"
        .Range("B5").Value = "West"
        
        .Range("C2").Value = "Fancy"
        .Range("C3").Value = "Fancy"
        .Range("C4").Value = "Tee"
        .Range("C5").Value = "Tee"
        
        .Range("D2").Value = 42.5
        .Range("D3").Value = 35.2
        .Range("D4").Value = 12.3
        .Range("D5").Value = 24.8
    End With
    
    ' Define data range
    Dim dataRange As Range
    Set dataRange = ThisWorkbook.ActiveSheet.Range("B1:D5")
    
    ' Add new worksheet for Pivot Table
    Dim pivotSheet As Worksheet
    Set pivotSheet = ThisWorkbook.Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotSheet.Range("A3"), _
        TableName:="SalesPivotTable")
    
    ' Add fields to Pivot Table
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Style").Orientation = xlRowField
        .AddDataField .PivotFields("Price"), "Sum of Price", xlSum
        .AddDataField .PivotFields("Price"), "Sum of Price 2", xlSum
    End With
    
    ' Inform user
    pivotSheet.Range("A16").Value = "Sum of Price will be moved soon"
    
    ' Schedule moving the data field after 5 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "MoveDataField"
End Sub

Sub MoveDataField()
    ' Move the data field to rows
    Dim pivotTable As PivotTable
    Set pivotTable = ThisWorkbook.Worksheets("PivotTableSheet").PivotTables("SalesPivotTable")
    
    With pivotTable
        .PivotFields("Sum of Price").Orientation = xlRowField
    End With
End Sub
```

```javascript
// JavaScript Code Equivalent for OnlyOffice API

// Set values in cells
var oWorksheet = Api.GetActiveSheet();

oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set header for Price

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

// Insert Pivot Table in new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to Pivot Table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data fields
pivotTable.AddDataField('Price');
pivotTable.AddDataField('Price');

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field for manipulation
var dataField = pivotTable.GetDataFields('Sum of Price');

// Inform user
pivotWorksheet.GetRange('A16').SetValue('Sum of Price will be moved soon');

// Schedule moving the data field after 5 seconds
setTimeout(function() {
	dataField.Move("Rows"); // Move data field to rows
}, 5000);
```