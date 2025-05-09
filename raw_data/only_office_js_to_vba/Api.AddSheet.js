## Description
**English:** This code creates a new worksheet named "New sheet".  
**Russian:** Этот код создает новый лист с именем "New sheet".

```javascript
// Create a new worksheet named "New sheet"
var oSheet = Api.AddSheet("New sheet");
```

```vba
' Create a new worksheet named "New sheet"
Dim oSheet As Worksheet
Set oSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
oSheet.Name = "New sheet"
```