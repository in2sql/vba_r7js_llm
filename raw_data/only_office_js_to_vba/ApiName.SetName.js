# Set and Rename Defined Range in Spreadsheet / Установка и переименование определённого диапазона в электронной таблице

This script sets values in cells A1 and B1, defines a named range for these cells, renames the defined name, and displays the new name in cell A3.
Этот скрипт устанавливает значения в ячейки A1 и B1, определяет именованный диапазон для этих ячеек, переименовывает определённое имя и отображает новое имя в ячейке A3.

```vba
' VBA Code to set values, define and rename a range, and display the new name

Sub SetAndRenameRange()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in A1 and B1
    oWorksheet.Range("A1").Value = "1"
    oWorksheet.Range("B1").Value = "2"
    
    ' Add a defined name for the range A1:B1
    ThisWorkbook.Names.Add Name:="name", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Retrieve the defined name
    Dim oDefName As Name
    Set oDefName = ThisWorkbook.Names("name")
    
    ' Rename the defined name to "new_name"
    oDefName.Name = "new_name"
    
    ' Retrieve the new defined name
    Dim oNewDefName As Name
    Set oNewDefName = ThisWorkbook.Names("new_name")
    
    ' Set value in A3 with the new name
    oWorksheet.Range("A3").Value = "The new name of the range: " & oNewDefName.Name
End Sub
```

```javascript
// JavaScript Code to set values, define and rename a range, and display the new name

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in A1 and B1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");

// Add a defined name for the range A1:B1
Api.AddDefName("name", "Sheet1!$A$1:$B$1");

// Retrieve the defined name
var oDefName = Api.GetDefName("name");

// Rename the defined name to "new_name"
oDefName.SetName("new_name");

// Retrieve the new defined name
var oNewDefName = Api.GetDefName("new_name");

// Set value in A3 with the new name
oWorksheet.GetRange("A3").SetValue("The new name of the range: " + oNewDefName.GetName());
```