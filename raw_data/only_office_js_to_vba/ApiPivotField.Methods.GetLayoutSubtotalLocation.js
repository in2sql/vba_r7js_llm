# Create and manipulate worksheet data, then insert a pivot table
# Создание и обработка данных на листе, затем вставка сводной таблицы

```javascript
// JavaScript OnlyOffice API code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set the value of cell B1 to 'Region'
oWorksheet.GetRange('C1').SetValue('Style');  // Set the value of cell C1 to 'Style'
oWorksheet.GetRange('D1').SetValue('Price');  // Set the value of cell D1 to 'Price'

// Set Region values
oWorksheet.GetRange('B2').SetValue('East');   // Set the value of cell B2 to 'East'
oWorksheet.GetRange('B3').SetValue('West');   // Set the value of cell B3 to 'West'
oWorksheet.GetRange('B4').SetValue('East');   // Set the value of cell B4 to 'East'
oWorksheet.GetRange('B5').SetValue('West');   // Set the value of cell B5 to 'West'

// Set Style values
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set the value of cell C2 to 'Fancy'
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set the value of cell C3 to 'Fancy'
oWorksheet.GetRange('C4').SetValue('Tee');    // Set the value of cell C4 to 'Tee'
oWorksheet.GetRange('C5').SetValue('Tee');    // Set the value of cell C5 to 'Tee'

// Set Price values
oWorksheet.GetRange('D2').SetValue(42.5);     // Set the value of cell D2 to 42.5
oWorksheet.GetRange('D3').SetValue(35.2);     // Set the value of cell D3 to 35.2
oWorksheet.GetRange('D4').SetValue(12.3);     // Set the value of cell D4 to 12.3
oWorksheet.GetRange('D5').SetValue(24.8);     // Set the value of cell D5 to 24.8

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'],
});

// Add data field to the pivot table
pivotTable.AddDataField('Price');

// Get the active worksheet (pivot table sheet)
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set description in the pivot worksheet
pivotWorksheet.GetRange('A14').SetValue('Region layout subtotal location'); // Set description text
pivotWorksheet.GetRange('B14').SetValue(pivotField.GetLayoutSubtotalLocation()); // Get and set subtotal location
```

```vba
' VBA code equivalent to the OnlyOffice API example

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set headers
    oWorksheet.Range("B1").Value = "Region" ' Set the value of cell B1 to "Region"
    oWorksheet.Range("C1").Value = "Style"  ' Set the value of cell C1 to "Style"
    oWorksheet.Range("D1").Value = "Price"  ' Set the value of cell D1 to "Price"
    
    ' Set Region values
    oWorksheet.Range("B2").Value = "East"    ' Set the value of cell B2 to "East"
    oWorksheet.Range("B3").Value = "West"    ' Set the value of cell B3 to "West"
    oWorksheet.Range("B4").Value = "East"    ' Set the value of cell B4 to "East"
    oWorksheet.Range("B5").Value = "West"    ' Set the value of cell B5 to "West"
    
    ' Set Style values
    oWorksheet.Range("C2").Value = "Fancy"   ' Set the value of cell C2 to "Fancy"
    oWorksheet.Range("C3").Value = "Fancy"   ' Set the value of cell C3 to "Fancy"
    oWorksheet.Range("C4").Value = "Tee"     ' Set the value of cell C4 to "Tee"
    oWorksheet.Range("C5").Value = "Tee"     ' Set the value of cell C5 to "Tee"
    
    ' Set Price values
    oWorksheet.Range("D2").Value = 42.5      ' Set the value of cell D2 to 42.5
    oWorksheet.Range("D3").Value = 35.2      ' Set the value of cell D3 to 35.2
    oWorksheet.Range("D4").Value = 12.3      ' Set the value of cell D4 to 12.3
    oWorksheet.Range("D5").Value = 24.8      ' Set the value of cell D5 to 24.8
    
    ' Define data range for pivot table
    Dim dataRef As Range
    Set dataRef = oWorksheet.Range("B1:D5")
    
    ' Add a new worksheet for the pivot table
    Dim pivotSheet As Worksheet
    Set pivotSheet = Worksheets.Add
    pivotSheet.Name = "PivotTableSheet"
    
    ' Create the pivot table
    Dim pivotTable As PivotTable
    Set pivotTable = pivotSheet.PivotTableWizard(TableDestination:=pivotSheet.Range("A1"), TableName:="PivotTable1", SourceData:=dataRef)
    
    ' Add row fields to the pivot table
    pivotTable.PivotFields("Region").Orientation = xlRowField
    pivotTable.PivotFields("Style").Orientation = xlRowField
    
    ' Add data field to the pivot table
    pivotTable.AddDataField pivotTable.PivotFields("Price"), "Sum of Price", xlSum
    
    ' Get the 'Region' pivot field
    Dim pivotField As PivotField
    Set pivotField = pivotTable.PivotFields("Region")
    
    ' Set description in the pivot worksheet
    pivotSheet.Range("A14").Value = "Region layout subtotal location" ' Set description text
    pivotSheet.Range("B14").Value = pivotField.Subtotals(1) ' Example: Get and set subtotal location
End Sub
```