# Code Description / Описание кода

**English:** This code sets specific values in cells of the active worksheet and creates a pivot table based on the defined data range.

**Russian:** Этот код устанавливает определенные значения в ячейки активного листа и создает сводную таблицу на основе заданного диапазона данных.

```vba
' Excel VBA code

Sub CreatePivotTable()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Set values in cells
    ws.Range("B1").Value = "Region"
    ws.Range("C1").Value = "Price"
    ws.Range("B2").Value = "East"
    ws.Range("B3").Value = "West"
    ws.Range("C2").Value = 42.5
    ws.Range("C3").Value = 35.2
    
    ' Define data range
    Dim dataRef As Range
    Set dataRef = ws.Range("B1:C3")
    
    ' Define pivot table location
    Dim pivotRef As Range
    Set pivotRef = ws.Range("A7")
    
    ' Create Pivot Cache
    Dim pivotCache As PivotCache
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRef)
    
    ' Create Pivot Table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotRef, _
        TableName:="PivotTable1")
    
    ' Optionally, set up pivot fields
    With pivotTable
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Price").Orientation = xlDataField
    End With
End Sub
```

```javascript
// OnlyOffice JS code

function createPivotTable() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set values in specific cells
    oWorksheet.GetRange('B1').SetValue('Region');
    oWorksheet.GetRange('C1').SetValue('Price');
    oWorksheet.GetRange('B2').SetValue('East');
    oWorksheet.GetRange('B3').SetValue('West');
    oWorksheet.GetRange('C2').SetValue(42.5);
    oWorksheet.GetRange('C3').SetValue(35.2);
    
    // Define the data range for the pivot table
    var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
    
    // Define the location where the pivot table will be inserted
    var pivotRef = oWorksheet.GetRange('A7');
    
    // Insert the pivot table into the worksheet
    var pivotTable = Api.InsertPivotExistingWorksheet(dataRef, pivotRef);
}

// Execute the function to create the pivot table
createPivotTable();
```