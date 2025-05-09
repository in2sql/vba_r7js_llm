## Description / Описание

**English:** This code demonstrates how to define a named range in a spreadsheet, set values in specific cells, and apply bold formatting to the range.

**Русский:** Этот код демонстрирует, как определить именованный диапазон в таблице, установить значения в определенные ячейки и применить жирное форматирование к диапазону.

---

### JavaScript (OnlyOffice API)

```javascript
// This example shows how to get the ApiRange object by its name.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set the value "1" in cell A1
oWorksheet.GetRange("B1").SetValue("2"); // Set the value "2" in cell B1
Api.AddDefName("numbers", "Sheet1!$A$1:$B$1"); // Define a named range "numbers" for cells A1:B1
var oDefName = Api.GetDefName("numbers"); // Get the defined named range "numbers"
var oRange = oDefName.GetRefersToRange(); // Get the range object referred to by the named range
oRange.SetBold(true); // Set the font to bold in the range
```

### Excel VBA

```vba
' This example shows how to get the Range object by its name.
Sub DefineNamedRange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    ws.Range("A1").Value = "1" ' Set the value "1" in cell A1
    ws.Range("B1").Value = "2" ' Set the value "2" in cell B1
    ThisWorkbook.Names.Add Name:="numbers", RefersTo:=ws.Range("A1:B1") ' Define a named range "numbers" for cells A1:B1
    With ThisWorkbook.Names("numbers").RefersToRange ' Get the range object referred to by the named range
        .Font.Bold = True ' Set the font to bold in the range
    End With
End Sub
```