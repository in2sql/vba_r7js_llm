---

**Description / Описание**

This code sets up data in specific cells of a worksheet, creates a pivot table based on that data, adds fields to the pivot table, and refreshes all pivot tables.

Этот код заполняет определенные ячейки листа данными, создает сводную таблицу на основе этих данных, добавляет поля в сводную таблицу и обновляет все сводные таблицы.

---

```javascript
// JavaScript OnlyOffice API code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in specific cells
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Price');  // Set header for Price
oWorksheet.GetRange('B2').SetValue('East');   // Set value East
oWorksheet.GetRange('B3').SetValue('West');   // Set value West
oWorksheet.GetRange('C2').SetValue(42.5);     // Set price 42.5
oWorksheet.GetRange('C3').SetValue(35.2);     // Set price 35.2

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");

// Insert a new pivot table on a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add Region as row field to the pivot table
Api.GetPivotByName(pivotTable.GetName()).AddFields({
	rows: 'Region',
});

// Add Price as data field to the pivot table
Api.GetPivotByName(pivotTable.GetName()).AddDataField('Price');

// Refresh all pivot tables
Api.RefreshAllPivots();
```

```vba
' Excel VBA equivalent code

Sub CreatePivotTable()
    ' Declare variables
    Dim oWorksheet As Worksheet
    Dim dataRef As Range
    Dim pivotTbl As PivotTable
    Dim pivotCache As PivotCache
    Dim pivotWS As Worksheet
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in specific cells
    oWorksheet.Range("B1").Value = "Region" ' Set header for Region
    oWorksheet.Range("C1").Value = "Price"  ' Set header for Price
    oWorksheet.Range("B2").Value = "East"    ' Set value East
    oWorksheet.Range("B3").Value = "West"    ' Set value West
    oWorksheet.Range("C2").Value = 42.5      ' Set price 42.5
    oWorksheet.Range("C3").Value = 35.2      ' Set price 35.2
    
    ' Define the data range for the pivot table
    Set dataRef = oWorksheet.Range("B1:C3")
    
    ' Create a Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRef)
    
    ' Add a new worksheet for the pivot table
    Set pivotWS = ThisWorkbook.Worksheets.Add
    pivotWS.Name = "PivotTableSheet"
    
    ' Create the Pivot Table
    Set pivotTbl = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWS.Range("A1"), _
        TableName:="PivotTable1")
    
    ' Add 'Region' as row field
    With pivotTbl.PivotFields("Region")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ' Add 'Price' as data field
    pivotTbl.AddDataField pivotTbl.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Refresh the pivot table
    pivotTbl.RefreshTable
End Sub
```