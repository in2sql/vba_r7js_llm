```plaintext
// This example sets values in cells A1 and B1, creates a named range 'numbers' referring to these cells, retrieves the range by its name, and sets the text to bold.
// Этот пример устанавливает значения в ячейки A1 и B1, создает именованный диапазон 'numbers', ссылающийся на эти ячейки, получает диапазон по его имени и устанавливает текст полужирным.
```

```vba
' This example sets values in cells A1 and B1, creates a named range "numbers" referring to these cells, retrieves the range by its name, and sets the text to bold.
Sub Example()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active sheet
    
    ws.Range("A1").Value = "1" ' Set value in A1
    ws.Range("B1").Value = "2" ' Set value in B1
    
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:=ws.Range("A1:B1") ' Add named range "numbers"
    
    Dim oRange As Range
    Set oRange = ThisWorkbook.Names("numbers").RefersToRange ' Get the range by name
    oRange.Font.Bold = True ' Set text to bold
End Sub
```

```javascript
// This example sets values in cells A1 and B1, creates a named range 'numbers' referring to these cells, retrieves the range by its name, and sets the text to bold.
// Этот пример устанавливает значения в ячейки A1 и B1, создает именованный диапазон 'numbers', ссылающийся на эти ячейки, получает диапазон по его имени и устанавливает текст полужирным.

var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
oWorksheet.GetRange("A1").SetValue("1"); // Set value in A1
oWorksheet.GetRange("B1").SetValue("2"); // Set value in B1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1"); // Add named range 'numbers'
var oDefName = Api.GetDefName("numbers"); // Get the defined name 'numbers'
var oRange = oDefName.GetRefersToRange(); // Get the range referred to by 'numbers'
oRange.SetBold(true); // Set text to bold
```