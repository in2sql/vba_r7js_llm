**Description / Описание:**

English: This example creates a new worksheet named "New sheet".

Russian: Этот пример создает новый рабочий лист с именем "New sheet".

```javascript
// This example creates a new worksheet.
var oSheet = Api.AddSheet("New sheet");
```

```vba
' Adds a new worksheet named "New sheet"
Dim oSheet As Worksheet
Set oSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
oSheet.Name = "New sheet"
```