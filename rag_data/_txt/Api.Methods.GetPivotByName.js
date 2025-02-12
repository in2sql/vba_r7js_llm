# Create and Populate a Worksheet with Data and a Pivot Table
# Создание и заполнение рабочей таблицы данными и сводной таблицей

## JavaScript Code / JavaScript Код

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in specific cells
oWorksheet.GetRange('B1').SetValue('Region'); // Set header for Region
oWorksheet.GetRange('C1').SetValue('Price');  // Set header for Price
oWorksheet.GetRange('B2').SetValue('East');   // Set value 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set value 'West' in cell B3
oWorksheet.GetRange('C2').SetValue(42.5);      // Set value 42.5 in cell C2
oWorksheet.GetRange('C3').SetValue(35.2);      // Set value 35.2 in cell C3

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add 'Region' as a row field in the pivot table
Api.GetPivotByName(pivotTable.GetName()).AddFields({
    rows: 'Region',
});

// Add 'Price' as a data field in the pivot table
Api.GetPivotByName(pivotTable.GetName()).AddDataField('Price');
```

## VBA Code / VBA Код

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set values in specific cells
oWorksheet.Range("B1").Value = "Region"   ' Set header for Region
oWorksheet.Range("C1").Value = "Price"    ' Set header for Price
oWorksheet.Range("B2").Value = "East"     ' Set value 'East' in cell B2
oWorksheet.Range("B3").Value = "West"     ' Set value 'West' in cell B3
oWorksheet.Range("C2").Value = 42.5        ' Set value 42.5 in cell C2
oWorksheet.Range("C3").Value = 35.2        ' Set value 35.2 in cell C3

' Define the data range for the pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:C3")

' Add a new worksheet for the pivot table
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add
pivotSheet.Name = "PivotTableSheet"

' Create a PivotCache
Dim pivotCache As PivotCache
Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRef)

' Create the PivotTable
Dim pivotTable As PivotTable
Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1")

' Add 'Region' as a row field
With pivotTable.PivotFields("Region")
    .Orientation = xlRowField
    .Position = 1
End With

' Add 'Price' as a data field
With pivotTable.PivotFields("Price")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Sum of Price"
End With
```