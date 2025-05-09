# Retrieve and List Theme Colors
# Получение и список цветовой темы

```javascript
// This example shows how to get a list of all the available theme colors for the spreadsheet.
// Этот пример показывает, как получить список всех доступных цветов темы для электронной таблицы.

var oWorksheet = Api.GetActiveSheet();
var themes = Api.GetThemesColors();
for (let i = 0; i < themes.length; ++i) {
    oWorksheet.GetRange("A" + (i + 1)).SetValue(themes[i]);
}
```

```vba
' This example shows how to get a list of all the available theme colors for the spreadsheet.
' Этот пример показывает, как получить список всех доступных цветов темы для электронной таблицы.

Sub ListThemeColors()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim themeColors As Variant
    themeColors = ActiveWorkbook.Theme.ThemeColorScheme.Colors
    
    Dim i As Integer
    For i = LBound(themeColors) To UBound(themeColors)
        ws.Range("A" & (i + 1)).Value = themeColors(i)
    Next i
End Sub
```