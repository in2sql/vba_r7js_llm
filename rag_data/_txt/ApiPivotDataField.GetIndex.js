**Description / Описание**

This script populates an OnlyOffice spreadsheet with region, style, and price data, creates a pivot table summarizing the data, and retrieves the index of the "Sum of Price" data field.

Этот скрипт заполняет электронную таблицу OnlyOffice данными о регионе, стиле и цене, создает сводную таблицу для суммирования данных и получает индекс поля данных "Сумма цены".

```javascript
// JavaScript OnlyOffice API code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');   // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');   // Set 'West' in cell B5

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set 'Tee' in cell C5

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);     // Set 42.5 in cell D2
oWorksheet.GetRange('D3').SetValue(35.2);     // Set 35.2 in cell D3
oWorksheet.GetRange('D4').SetValue(12.3);     // Set 12.3 in cell D4
oWorksheet.GetRange('D5').SetValue(24.8);     // Set 24.8 in cell D5

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
    rows: ['Region', 'Style'],
});

// Add data fields to the pivot table
pivotTable.AddDataField('Price');
pivotTable.AddDataField('Price');

// Get the active pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the data field for 'Sum of Price'
var dataField = pivotTable.GetDataFields('Sum of Price');

// Set labels and values in the pivot worksheet
pivotWorksheet.GetRange('A15').SetValue('Sum of Price index:'); // Label in A15
pivotWorksheet.GetRange('B15').SetValue(dataField.GetIndex());    // Index value in B15
```

```vba
' VBA Equivalent Code

Sub CreatePivotTable()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Set headers
    oWorksheet.GetRange("B1").SetValue "Region" ' Set 'Region' in cell B1
    oWorksheet.GetRange("C1").SetValue "Style"  ' Set 'Style' in cell C1
    oWorksheet.GetRange("D1").SetValue "Price"  ' Set 'Price' in cell D1
    
    ' Populate Region data
    oWorksheet.GetRange("B2").SetValue "East"   ' Set 'East' in cell B2
    oWorksheet.GetRange("B3").SetValue "West"   ' Set 'West' in cell B3
    oWorksheet.GetRange("B4").SetValue "East"   ' Set 'East' in cell B4
    oWorksheet.GetRange("B5").SetValue "West"   ' Set 'West' in cell B5
    
    ' Populate Style data
    oWorksheet.GetRange("C2").SetValue "Fancy"  ' Set 'Fancy' in cell C2
    oWorksheet.GetRange("C3").SetValue "Fancy"  ' Set 'Fancy' in cell C3
    oWorksheet.GetRange("C4").SetValue "Tee"    ' Set 'Tee' in cell C4
    oWorksheet.GetRange("C5").SetValue "Tee"    ' Set 'Tee' in cell C5
    
    ' Populate Price data
    oWorksheet.GetRange("D2").SetValue 42.5     ' Set 42.5 in cell D2
    oWorksheet.GetRange("D3").SetValue 35.2     ' Set 35.2 in cell D3
    oWorksheet.GetRange("D4").SetValue 12.3     ' Set 12.3 in cell D4
    oWorksheet.GetRange("D5").SetValue 24.8     ' Set 24.8 in cell D5
    
    ' Define the data range for the pivot table
    Dim dataRef As Object
    Set dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5")
    
    ' Insert a new pivot table worksheet
    Dim pivotTable As Object
    Set pivotTable = Api.InsertPivotNewWorksheet(dataRef)
    
    ' Add row fields to the pivot table
    pivotTable.AddFields Array("Region", "Style"), , , False
    
    ' Add data fields to the pivot table
    pivotTable.AddDataField "Price"
    pivotTable.AddDataField "Price"
    
    ' Get the active pivot worksheet
    Dim pivotWorksheet As Object
    Set pivotWorksheet = Api.GetActiveSheet()
    
    ' Get the data field for 'Sum of Price'
    Dim dataField As Object
    Set dataField = pivotTable.GetDataFields("Sum of Price")
    
    ' Set labels and values in the pivot worksheet
    pivotWorksheet.GetRange("A15").SetValue "Sum of Price index:" ' Label in A15
    pivotWorksheet.GetRange("B15").SetValue dataField.GetIndex()   ' Index value in B15
End Sub
```