### Description / Описание
**English:**  
This example demonstrates how to set values in cells A1 and B1, define a named range "numbers" for these cells, retrieve the defined name, and display the name in cell A3.

**Русский:**  
Этот пример демонстрирует, как установить значения в ячейках A1 и B1, определить именованный диапазон "numbers" для этих ячеек, получить определенное имя и отобразить имя в ячейке A3.

```vba
' VBA Code
Sub DefineAndRetrieveNamedRange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in A1 and B1
    ws.Range("A1").Value = "1"
    ws.Range("B1").Value = "2"
    
    ' Add a named range "numbers" referring to A1:B1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Retrieve the named range "numbers"
    Dim defName As Name
    Set defName = ThisWorkbook.Names("numbers")
    
    ' Set value in A3 to display the defined name
    ws.Range("A3").Value = "DefName: " & defName.Name
End Sub
```

```javascript
// JavaScript Code
// This example shows how to get the ApiName object by the range name.
var oWorksheet = Api.GetActiveSheet();

// Set values in A1 and B1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");

// Add a named range "numbers" referring to A1:B1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");

// Retrieve the named range "numbers"
var oDefName = Api.GetDefName("numbers");

// Set value in A3 to display the defined name
oWorksheet.GetRange("A3").SetValue("DefName: " + oDefName.GetName());
```