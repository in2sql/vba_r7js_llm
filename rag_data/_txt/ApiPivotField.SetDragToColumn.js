**Description / Описание:**
This script sets up a worksheet with headers and data, creates a pivot table based on that data, configures the pivot table fields, and adds specific labels to the pivot worksheet.  
Этот скрипт настраивает рабочий лист с заголовками и данными, создает сводную таблицу на основе этих данных, настраивает поля сводной таблицы и добавляет определенные метки на сводный рабочий лист.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set 'Region' in cell B1
oWorksheet.GetRange('C1').SetValue('Style');  // Set 'Style' in cell C1
oWorksheet.GetRange('D1').SetValue('Price');  // Set 'Price' in cell D1

// Set Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set 'East' in cell B2
oWorksheet.GetRange('B3').SetValue('West');   // Set 'West' in cell B3
oWorksheet.GetRange('B4').SetValue('East');   // Set 'East' in cell B4
oWorksheet.GetRange('B5').SetValue('West');   // Set 'West' in cell B5

// Set Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set 'Fancy' in cell C2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set 'Fancy' in cell C3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set 'Tee' in cell C4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set 'Tee' in cell C5

// Set Price data
oWorksheet.GetRange('D2').SetValue(42.5);      // Set 42.5 in cell D2
oWorksheet.GetRange('D3').SetValue(35.2);      // Set 35.2 in cell D3
oWorksheet.GetRange('D4').SetValue(12.3);      // Set 12.3 in cell D4
oWorksheet.GetRange('D5').SetValue(24.8);      // Set 24.8 in cell D5

// Define data range for pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add fields to pivot table
pivotTable.AddFields({
	columns: ['Style'], // Add 'Style' to columns
	rows: 'Region',      // Add 'Region' to rows
});

// Add data field to pivot table
pivotTable.AddDataField('Price'); // Add 'Price' as data field

// Get the pivot worksheet
var pivotWorksheet = Api.GetActiveSheet();

// Get the 'Region' pivot field
var pivotField = pivotTable.GetPivotFields('Region');

// Set 'Drag to column' property to false for 'Region' field
pivotField.SetDragToColumn(false);

// Add labels to pivot worksheet
pivotWorksheet.GetRange('A13').SetValue('Drag to column');                      // Set label in A13
pivotWorksheet.GetRange('B13').SetValue(pivotField.GetDragToColumn());          // Set value in B13
pivotWorksheet.GetRange('A14').SetValue('Try drag Region to columns!');        // Set instruction in A14
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set headers
oWorksheet.Range("B1").Value = "Region" ' Set "Region" in cell B1
oWorksheet.Range("C1").Value = "Style"  ' Set "Style" in cell C1
oWorksheet.Range("D1").Value = "Price"  ' Set "Price" in cell D1

' Set Region data
oWorksheet.Range("B2").Value = "East"   ' Set "East" in cell B2
oWorksheet.Range("B3").Value = "West"   ' Set "West" in cell B3
oWorksheet.Range("B4").Value = "East"   ' Set "East" in cell B4
oWorksheet.Range("B5").Value = "West"   ' Set "West" in cell B5

' Set Style data
oWorksheet.Range("C2").Value = "Fancy"  ' Set "Fancy" in cell C2
oWorksheet.Range("C3").Value = "Fancy"  ' Set "Fancy" in cell C3
oWorksheet.Range("C4").Value = "Tee"    ' Set "Tee" in cell C4
oWorksheet.Range("C5").Value = "Tee"    ' Set "Tee" in cell C5

' Set Price data
oWorksheet.Range("D2").Value = 42.5     ' Set 42.5 in cell D2
oWorksheet.Range("D3").Value = 35.2     ' Set 35.2 in cell D3
oWorksheet.Range("D4").Value = 12.3     ' Set 12.3 in cell D4
oWorksheet.Range("D5").Value = 24.8     ' Set 24.8 in cell D5

' Define data range for pivot table
Dim dataRef As Range
Set dataRef = oWorksheet.Range("B1:D5")

' Insert pivot table in a new worksheet
Dim pivotTable As PivotTable
Dim pivotSheet As Worksheet
Set pivotSheet = ThisWorkbook.Worksheets.Add
Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, SourceData:=dataRef)

' Add fields to pivot table
With pivotTable
    .PivotFields("Style").Orientation = xlColumnField    ' Add "Style" to columns
    .PivotFields("Region").Orientation = xlRowField      ' Add "Region" to rows
    .AddDataField .PivotFields("Price"), "Sum of Price", xlSum ' Add "Price" as data field
End With

' Get the 'Region' pivot field
Dim pivotField As PivotField
Set pivotField = pivotTable.PivotFields("Region")

' Set 'Drag to column' property to false for 'Region' field
' VBA does not have a direct equivalent for SetDragToColumn, so this step may require additional handling based on specific requirements

' Add labels to pivot worksheet
pivotSheet.Range("A13").Value = "Drag to column"                       ' Set label in A13
pivotSheet.Range("B13").Value = pivotField.Orientation = xlColumnField ' Set value in B13
pivotSheet.Range("A14").Value = "Try drag Region to columns!"         ' Set instruction in A14
```