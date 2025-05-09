### Description / Описание

This code retrieves the active worksheet, obtains the worksheet functions, and sets the value of cell A1 to the ASCII code of the string "text".

Этот код получает активный лист, получает функции листа и устанавливает значение ячейки A1 в ASCII код строки "text".

```vba
' VBA Code

' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set the value of cell A1 to the ASCII code of "text"
oWorksheet.Range("A1").Value = Asc("text")
```

```javascript
// OnlyOffice JavaScript Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get worksheet functions
var oFunction = Api.GetWorksheetFunction();

// Set the value of cell A1 to the result of ASC("text")
oWorksheet.GetRange("A1").SetValue(oFunction.ASC("text"));
```