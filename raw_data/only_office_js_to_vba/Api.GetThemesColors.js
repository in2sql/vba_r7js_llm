# Retrieve and Populate Spreadsheet Theme Colors
# Получение и заполнение цветовых тем таблицы

```vba
' VBA code to retrieve all available theme colors and populate them in column A

Sub PopulateThemeColors()
    Dim oWorksheet As Worksheet
    Dim themes As Variant
    Dim i As Integer
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Get the theme colors from the active workbook
    themes = ActiveWorkbook.Theme.ThemeColorScheme.Colors
    
    ' Loop through the theme colors and set each color in column A
    For i = LBound(themes) To UBound(themes)
        oWorksheet.Range("A" & (i + 1)).Value = themes(i)
    Next i
End Sub
```

```javascript
// This example shows how to get a list of all the available theme colors for the spreadsheet.
// This code retrieves all available theme colors and populates them into column A of the active worksheet.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var themes = Api.GetThemesColors(); // Retrieve theme colors
for (let i = 0; i < themes.length; ++i) {
    oWorksheet.GetRange("A" + (i + 1)).SetValue(themes[i]); // Set each theme color in column A
}
```