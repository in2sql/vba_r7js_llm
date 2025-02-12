### Description
**English**: This code demonstrates how to retrieve an `ApiName` object using a range name, set values in specific cells, add a defined name, retrieve it, and display its name in another cell.

**Russian**: Этот код демонстрирует, как получить объект `ApiName` по имени диапазона, установить значения в определенные ячейки, добавить определенное имя, получить его и отобразить его имя в другой ячейке.

```vba
' VBA code equivalent
' This macro retrieves an ApiName object by range name, sets cell values, adds a defined name, and displays the name.

Sub Example()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in cells A1 and B1
    oWorksheet.Range("A1").Value = "1"
    oWorksheet.Range("B1").Value = "2"
    
    ' Add a defined name "numbers" referring to Sheet1!$A$1:$B$1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Get the defined name "numbers"
    Dim oDefName As Name
    Set oDefName = ThisWorkbook.Names("numbers")
    
    ' Set value in A3 with the name of the defined name
    oWorksheet.Range("A3").Value = "DefName: " & oDefName.Name
End Sub
```

```javascript
// JavaScript code equivalent
// This example shows how to get the ApiName object by the range name, set cell values, add a defined name, and display the name.

var oWorksheet = Api.GetActiveSheet();

// Set values in cells A1 and B1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");

// Add a defined name "numbers" referring to Sheet1!$A$1:$B$1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");

// Get the defined name "numbers"
var oDefName = Api.GetDefName("numbers");

// Set value in A3 with the name of the defined name
oWorksheet.GetRange("A3").SetValue("DefName: " + oDefName.GetName());
```