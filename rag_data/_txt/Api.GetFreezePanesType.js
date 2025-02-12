## Description / Описание

**English:**  
This code freezes the first column and writes the freeze pane type into cells A1 and B1.

**Русский:**  
Этот код замораживает первый столбец и записывает тип замороженной панели в ячейки A1 и B1.

---

### VBA Code

```vba
' This VBA macro freezes the first column and writes the freeze pane type into cells A1 and B1
Sub FreezeFirstColumnAndSetType()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Freeze the first column
    ws.Columns("B").Select
    ActiveWindow.FreezePanes = True
    
    ' Set values in A1 and B1
    ws.Range("A1").Value = "Type: "
    ws.Range("B1").Value = "Column"
End Sub
```

### OnlyOffice JavaScript Code

```javascript
// This example freezes the first column and writes the freeze pane type into cells A1 and B1
Api.SetFreezePanesType('column');
var oWorksheet = Api.GetActiveSheet();

// Set the value in cell A1
oWorksheet.GetRange("A1").SetValue("Type: ");

// Get the freeze pane type and set it in cell B1
oWorksheet.GetRange("B1").SetValue(Api.GetFreezePanesType());
```