### Description / Описание

**English:**  
This code sets values in cells A1 and B1, defines a named range "numbers" referring to these cells, retrieves the defined name, and sets the value in cell A3 to display the name.

**Russian:**  
Этот код устанавливает значения в ячейки A1 и B1, определяет именованный диапазон "numbers", ссылающийся на эти ячейки, извлекает определенное имя и устанавливает значение в ячейку A3 для отображения имени.

```vba
' VBA Code Equivalent

Sub DefineAndRetrieveNamedRange()
    Dim oWorksheet As Worksheet
    Dim oDefName As Name
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Set value "2" in cell B1
    oWorksheet.Range("B1").Value = "2"
    
    ' Add a defined name "numbers" referring to A1:B1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:=oWorksheet.Range("A1:B1")
    
    ' Get the defined name "numbers"
    Set oDefName = ThisWorkbook.Names("numbers")
    
    ' Set value in cell A3 to display the defined name
    oWorksheet.Range("A3").Value = "Name: " & oDefName.Name
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example sets values in A1 and B1, defines a named range "numbers",
// retrieves the defined name, and sets the value in A3 to display the name.

var oWorksheet = Api.GetActiveSheet();

// Set value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Set value "2" in cell B1
oWorksheet.GetRange("B1").SetValue("2");

// Add a defined name "numbers" referring to A1:B1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");

// Get the defined name "numbers"
var oDefName = Api.GetDefName("numbers");

// Set value in cell A3 to display the defined name
oWorksheet.GetRange("A3").SetValue("Name: " + oDefName.GetName());
```