**Description / Описание**

English: This code sets values in cells A1 and B1, adds a defined name "numbers" for the range A1:B1, retrieves and deletes this defined name, and writes a message in cell A3 indicating the deletion.

Russian: Этот код устанавливает значения в ячейки A1 и B1, добавляет определённое имя "numbers" для диапазона A1:B1, получает и удаляет это определённое имя, а затем записывает сообщение в ячейку A3, указывая на удаление.

```vba
' VBA code to delete a defined name and update cell A3 accordingly

Sub DeleteDefNameExample()
    Dim oDefName As Name
    Dim oWorksheet As Worksheet

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"

    ' Set value "2" in cell B1
    oWorksheet.Range("B1").Value = "2"

    ' Add defined name "numbers" for range A1:B1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:="=Sheet1!$A$1:$B$1"

    ' Get the defined name "numbers"
    Set oDefName = ThisWorkbook.Names("numbers")

    ' Delete the defined name "numbers"
    oDefName.Delete

    ' Set value in cell A3 indicating the deletion
    oWorksheet.Range("A3").Value = "The name 'numbers' of the range A1:B1 was deleted."
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to delete the DefName object

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Set value "2" in cell B1
oWorksheet.GetRange("B1").SetValue("2");

// Add defined name "numbers" for range A1:B1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");

// Get the defined name "numbers"
var oDefName = Api.GetDefName("numbers");

// Delete the defined name "numbers"
oDefName.Delete();

// Set value in cell A3 indicating the deletion
oWorksheet.GetRange("A3").SetValue("The name 'numbers' of the range A1:B1 was deleted.");
```