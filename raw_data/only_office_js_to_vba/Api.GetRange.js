**Description:**
- **English:** This code retrieves the active worksheet, selects the range A1:C1, sets its fill color to a specific RGB value, and writes a message in cell A3 indicating that the color has been set for the background of cells A1:C1.
- **Russian:** Этот код получает активный рабочий лист, выбирает диапазон A1:C1, устанавливает его цвет заливки на определенное значение RGB и записывает сообщение в ячейку A3, указывая, что цвет был установлен для фона ячеек A1:C1.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range A1:C1
var oRange = Api.GetRange("A1:C1");

// Set the fill color of the range to RGB(255, 213, 191)
oRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));

// Set the value of cell A3 with a message
oWorksheet.GetRange("A3").SetValue("The color was set to the background of cells A1:C1.");
```

```vba
' Получение активного рабочего листа
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet()

' Получение диапазона A1:C1
Dim oRange As Object
Set oRange = Api.GetRange("A1:C1")

' Установка цвета заливки диапазона на RGB(255, 213, 191)
oRange.SetFillColor Api.CreateColorFromRGB(255, 213, 191)

' Установка значения ячейки A3 с сообщением
oWorksheet.GetRange("A3").SetValue "The color was set to the background of cells A1:C1."
```