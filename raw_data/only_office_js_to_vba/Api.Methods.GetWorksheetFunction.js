# Sets the value of cell A1 to the ASCII code of the string "text"
# Устанавливает значение ячейки A1 на ASCII-код строки "text"

## Excel VBA:

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set the value of cell A1 to the ASCII code of "text"
oWorksheet.Range("A1").Value = Asc("text")
```

## OnlyOffice JavaScript:

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get worksheet functions
var oFunction = Api.GetWorksheetFunction();

// Set the value of cell A1 to the ASCII code of "text"
oWorksheet.GetRange("A1").SetValue(oFunction.ASC("text"));
```