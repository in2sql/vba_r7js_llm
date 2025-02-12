**Description:**
- **English:** This example demonstrates how to retrieve the full name of the currently opened file and set it in cell B1 of the active worksheet.
- **Russian:** Этот пример демонстрирует, как получить полное имя текущего открытого файла и установить его в ячейку B1 активного листа.

```vba
' VBA code to get the full name of the currently opened workbook and set it in cell B1 of the active sheet
Sub SetFileName()
    Dim sName As String
    sName = ThisWorkbook.FullName ' Get the full name of the current workbook
    With ActiveSheet
        .Range("B1").Value = "File name: " & sName ' Set the value in B1
    End With
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to get the full name of the currently opened file and set it in cell B1 of the active sheet
var oWorksheet = Api.GetActiveSheet();
var sName = Api.GetFullName();
oWorksheet.GetRange("B1").SetValue("File name: " + sName);
```