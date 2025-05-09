## Description / Описание

**English:**  
This example deletes the defined name "numbers" associated with the range A1:B1 in the active worksheet.

**Russian:**  
Этот пример удаляет определённое имя "numbers", связанное с диапазоном A1:B1 на активном листе.

---

### Excel VBA Code

```vba
' This example deletes the DefName object.
Sub DeleteDefName()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value "1" in A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Set value "2" in B1
    oWorksheet.Range("B1").Value = "2"
    
    ' Add defined name "numbers" referring to A1:B1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Delete the defined name "numbers"
    ThisWorkbook.Names("numbers").Delete
    
    ' Set message in A3
    oWorksheet.Range("A3").Value = "The name 'numbers' of the range A1:B1 was deleted."
End Sub
```

---

### OnlyOffice JS Code

```javascript
// This example deletes the DefName object.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1"); // Set value "1" in A1
oWorksheet.GetRange("B1").SetValue("2"); // Set value "2" in B1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1"); // Add defined name "numbers" for range A1:B1
var oDefName = Api.GetDefName("numbers");
oDefName.Delete(); // Delete the defined name "numbers"
oWorksheet.GetRange("A3").SetValue("The name 'numbers' of the range A1:B1 was deleted."); // Set message in A3
```