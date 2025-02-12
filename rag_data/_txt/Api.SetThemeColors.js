---

**Description / Описание**

**English:**  
This script sets the theme colors to the current spreadsheet by retrieving the theme colors, writing them into column A, setting a specific theme color, and adding a descriptive message in cell C3.

**Русский:**  
Этот скрипт устанавливает цветовую тему текущей таблицы, извлекая цвета темы, записывая их в столбец A, устанавливая конкретный цвет темы и добавляя описательное сообщение в ячейку C3.

---

### Excel VBA Code

```vba
' This VBA script sets the theme colors to the current spreadsheet by retrieving the theme colors,
' writing them into column A, setting a specific theme color, and adding a descriptive message in cell C3.

Sub SetThemeColors()
    Dim oWorksheet As Worksheet
    Dim themes As Variant
    Dim i As Integer
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Retrieve the theme colors
    themes = ActiveWorkbook.Theme.ThemeColorScheme.Colors
    
    ' Write each theme color into column A
    For i = LBound(themes) To UBound(themes)
        oWorksheet.Range("A" & (i + 1)).Interior.Color = themes(i + 1)
    Next i
    
    ' Set the theme colors to the fourth theme color (equivalent to themes[3] in JS)
    ActiveWorkbook.Theme.ThemeColorScheme.Colors(msoThemeColorAccent4) = themes(4)
    
    ' Add a descriptive message in cell C3
    oWorksheet.Range("C3").Value = "The 'Apex' theme colors were set to the current spreadsheet."
End Sub
```

---

### OnlyOffice JavaScript Code

```javascript
// This example sets the theme colors to the current spreadsheet.

function setThemeColors() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Retrieve the theme colors
    var themes = Api.GetThemesColors();
    
    // Write each theme color into column A
    for (let i = 0; i < themes.length; ++i) {
        oWorksheet.GetRange("A" + (i + 1)).SetValue(themes[i]);
    }
    
    // Set the theme colors to the fourth theme color
    Api.SetThemeColors(themes[3]);
    
    // Add a descriptive message in cell C3
    oWorksheet.GetRange("C3").SetValue("The 'Apex' theme colors were set to the current spreadsheet.");
}
```

---