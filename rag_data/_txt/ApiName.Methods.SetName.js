**English:** This VBA and JavaScript code examples replicate the functionality of setting values in cells, defining a named range, renaming it, and displaying the new name in a specific cell.

**Russian:** Эти примеры кода на VBA и JavaScript воспроизводят функциональность установки значений в ячейки, определения именованного диапазона, его переименования и отображения нового имени в конкретной ячейке.

```vba
' VBA Code to set values, define a named range, rename it, and display the new name
Sub ManageNamedRange()
    Dim oWorksheet As Worksheet
    Dim oDefName As Name
    Dim oNewDefName As Name
    
    ' Get the active sheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in cells A1 and B1
    oWorksheet.Range("A1").Value = "1"
    oWorksheet.Range("B1").Value = "2"
    
    ' Add a defined name "name" referring to A1:B1
    ThisWorkbook.Names.Add Name:="name", RefersTo:="=" & oWorksheet.Name & "!$A$1:$B$1"
    
    ' Get the defined name "name"
    Set oDefName = ThisWorkbook.Names("name")
    
    ' Rename the defined name to "new_name"
    oDefName.Name = "new_name"
    
    ' Get the new defined name "new_name"
    Set oNewDefName = ThisWorkbook.Names("new_name")
    
    ' Set value in cell A3 with the new name
    oWorksheet.Range("A3").Value = "The new name of the range: " & oNewDefName.Name
End Sub
```

```javascript
// JavaScript Code to set values, define a named range, rename it, and display the new name

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells A1 and B1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");

// Add a defined name "name" referring to A1:B1
Api.AddDefName("name", "Sheet1!$A$1:$B$1");

// Get the defined name "name"
var oDefName = Api.GetDefName("name");

// Rename the defined name to "new_name"
oDefName.SetName("new_name");

// Get the new defined name "new_name"
var oNewDefName = Api.GetDefName("new_name");

// Set value in cell A3 with the new name
oWorksheet.GetRange("A3").SetValue("The new name of the range: " + oNewDefName.GetName());
```