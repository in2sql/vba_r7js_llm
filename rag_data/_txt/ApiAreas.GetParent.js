# Description
This code retrieves the parent object of a specified range in an Excel sheet, sets values in certain cells, and displays the parent and its type.
Этот код получает родительский объект указанного диапазона на листе Excel, устанавливает значения в определённых ячейках и отображает родителя и его тип.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1:D1
var oRange = oWorksheet.GetRange("B1:D1");

// Set the value of the range to "1"
oRange.SetValue("1");

// Select the range
oRange.Select();

// Get the areas of the range
var oAreas = oRange.GetAreas();

// Get the parent of the areas
var oParent = oAreas.GetParent();

// Get the class type of the parent
var sType = oParent.GetClassType();

// Set the value of cell A4
oRange = oWorksheet.GetRange('A4');
oRange.SetValue("The areas parent: ");

// Auto fit the column
oRange.AutoFit(false, true);

// Paste the parent object into cell B4
oWorksheet.GetRange('B4').Paste(oParent);

// Set the value of cell A5
oRange = oWorksheet.GetRange('A5');
oRange.SetValue("The type of the areas parent: ");

// Auto fit the column
oRange.AutoFit(false, true);

// Set the class type in cell B5
oWorksheet.GetRange('B5').SetValue(sType);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Get the range B1:D1
Dim oRange As Range
Set oRange = oWorksheet.Range("B1:D1")

' Set the value of the range to "1"
oRange.Value = "1"

' Select the range
oRange.Select

' Get the areas of the range
Dim oAreas As Areas
Set oAreas = oRange.Areas

' Get the parent of the areas
Dim oParent As Object
Set oParent = oAreas.Parent

' Get the class type of the parent
Dim sType As String
sType = TypeName(oParent)

' Set the value of cell A4
Set oRange = oWorksheet.Range("A4")
oRange.Value = "The areas parent: "

' Auto fit the column
oRange.EntireColumn.AutoFit

' Paste the parent object into cell B4
oWorksheet.Range("B4").Value = oParent.Name

' Set the value of cell A5
Set oRange = oWorksheet.Range("A5")
oRange.Value = "The type of the areas parent: "

' Auto fit the column
oRange.EntireColumn.AutoFit

' Set the class type in cell B5
oWorksheet.Range("B5").Value = sType
```