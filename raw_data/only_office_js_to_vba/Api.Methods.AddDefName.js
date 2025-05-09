# Description / Описание

This code adds a new name to a range of cells.  
Этот код добавляет новое имя к диапазону ячеек.

```vba
' VBA code to add a new name to a range of cells

Sub AddNamedRange()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value of cell A1
    oWorksheet.Range("A1").Value = 1
    
    ' Set value of cell B1
    oWorksheet.Range("B1").Value = 2
    
    ' Add defined name 'numbers' referring to range A1:B1 on Sheet1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Set value of cell A3 with a message
    oWorksheet.Range("A3").Value = "We defined a name 'numbers' for a range of cells A1:B1."
End Sub
```

```javascript
// JavaScript code to add a new name to a range of cells

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value of cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Set value of cell B1
oWorksheet.GetRange("B1").SetValue("2");

// Add defined name 'numbers' referring to range A1:B1 on Sheet1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");

// Set value of cell A3 with a message
oWorksheet.GetRange("A3").SetValue("We defined a name 'numbers' for a range of cells A1:B1.");
```