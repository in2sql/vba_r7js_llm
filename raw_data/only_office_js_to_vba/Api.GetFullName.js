### Description / Описание

**English:**  
This code retrieves the full name of the currently opened Excel file and sets the value of cell B1 in the active sheet to display the file name.

**Russian:**  
Этот код получает полное имя текущего открытого файла Excel и устанавливает значение ячейки B1 на активном листе для отображения имени файла.

```vba
' VBA Code
' Get the active sheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the full name of the workbook
Dim sName As String
sName = ThisWorkbook.FullName

' Set cell B1 to "File name: " followed by the full name
oWorksheet.Range("B1").Value = "File name: " & sName
```

```javascript
// OnlyOffice JS Code
// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Get the full name of the workbook
var sName = Api.GetFullName();

// Set cell B1 to "File name: " followed by the full name
oWorksheet.GetRange("B1").SetValue("File name: " + sName);
```