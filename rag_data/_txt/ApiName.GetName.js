**Description / Описание**

This script demonstrates how to interact with the active worksheet by setting values in specific cells, defining a named range, retrieving that name, and displaying it in another cell.

Этот скрипт демонстрирует, как взаимодействовать с активным листом, устанавливая значения в определенные ячейки, определяя именованный диапазон, получая это имя и отображая его в другой ячейке.

```vba
' VBA code equivalent to OnlyOffice JS example

Sub Example()
    Dim oWorksheet As Worksheet
    Dim oDefName As Name
    
    ' Get the active sheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Set value "2" in cell B1
    oWorksheet.Range("B1").Value = "2"
    
    ' Add a defined name "numbers" referring to A1:B1 in Sheet1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Get the defined name "numbers"
    Set oDefName = ThisWorkbook.Names("numbers")
    
    ' Set value in cell A3 with the name of the defined name
    oWorksheet.Range("A3").Value = "Name: " & oDefName.Name
End Sub
```

```javascript
// OnlyOffice JS code equivalent to VBA example

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Set value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Set value "2" in cell B1
oWorksheet.GetRange("B1").SetValue("2");

// Add a defined name "numbers" referring to A1:B1 in Sheet1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");

// Get the defined name "numbers"
var oDefName = Api.GetDefName("numbers");

// Set value in cell A3 with the name of the defined name
oWorksheet.GetRange("A3").SetValue("Name: " + oDefName.GetName());
```