**Description:**
This script initializes a worksheet by setting up headers and populating data for regions, styles, and prices. It then creates a pivot table based on this data, organizing it by region and style, and adding price and region as data fields.

**Описание:**
Этот скрипт инициализирует лист, устанавливая заголовки и заполняя данные о регионах, стилях и ценах. Затем он создает сводную таблицу на основе этих данных, организуя ее по регионам и стилям, и добавляет цену и регион в качестве полей данных.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set headers
oWorksheet.GetRange('B1').SetValue('Region'); // Set the header for Region
oWorksheet.GetRange('C1').SetValue('Style');  // Set the header for Style
oWorksheet.GetRange('D1').SetValue('Price');  // Set the header for Price

// Populate Region data
oWorksheet.GetRange('B2').SetValue('East');   // Set Region for row 2
oWorksheet.GetRange('B3').SetValue('West');   // Set Region for row 3
oWorksheet.GetRange('B4').SetValue('East');   // Set Region for row 4
oWorksheet.GetRange('B5').SetValue('West');   // Set Region for row 5

// Populate Style data
oWorksheet.GetRange('C2').SetValue('Fancy');  // Set Style for row 2
oWorksheet.GetRange('C3').SetValue('Fancy');  // Set Style for row 3
oWorksheet.GetRange('C4').SetValue('Tee');    // Set Style for row 4
oWorksheet.GetRange('C5').SetValue('Tee');    // Set Style for row 5

// Populate Price data
oWorksheet.GetRange('D2').SetValue(42.5);     // Set Price for row 2
oWorksheet.GetRange('D3').SetValue(35.2);     // Set Price for row 3
oWorksheet.GetRange('D4').SetValue(12.3);     // Set Price for row 4
oWorksheet.GetRange('D5').SetValue(24.8);     // Set Price for row 5

// Define the data range for the pivot table
var dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");

// Insert a new pivot table in a new worksheet
var pivotTable = Api.InsertPivotNewWorksheet(dataRef);

// Add row fields to the pivot table
pivotTable.AddFields({
	rows: ['Region', 'Style'], // Add Region and Style as row fields
});

// Add data field to the pivot table
pivotTable.AddDataField('Price'); // Add Price as a data field

// Get the active sheet where the pivot table is located
var pivotWorksheet = Api.GetActiveSheet();

// Get the pivot field for Style
var pivotField = pivotTable.GetPivotFields('Style');

// Add Region as a data field in the pivot table
pivotField.GetTable().AddDataField('Region'); 
```

```vba
' Получить активный лист
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet()

' Установить заголовки
oWorksheet.GetRange("B1").SetValue "Region" ' Установить заголовок для Region
oWorksheet.GetRange("C1").SetValue "Style"  ' Установить заголовок для Style
oWorksheet.GetRange("D1").SetValue "Price"  ' Установить заголовок для Price

' Заполнить данные Region
oWorksheet.GetRange("B2").SetValue "East"   ' Установить Region для строки 2
oWorksheet.GetRange("B3").SetValue "West"   ' Установить Region для строки 3
oWorksheet.GetRange("B4").SetValue "East"   ' Установить Region для строки 4
oWorksheet.GetRange("B5").SetValue "West"   ' Установить Region для строки 5

' Заполнить данные Style
oWorksheet.GetRange("C2").SetValue "Fancy"  ' Установить Style для строки 2
oWorksheet.GetRange("C3").SetValue "Fancy"  ' Установить Style для строки 3
oWorksheet.GetRange("C4").SetValue "Tee"    ' Установить Style для строки 4
oWorksheet.GetRange("C5").SetValue "Tee"    ' Установить Style для строки 5

' Заполнить данные Price
oWorksheet.GetRange("D2").SetValue 42.5      ' Установить Price для строки 2
oWorksheet.GetRange("D3").SetValue 35.2      ' Установить Price для строки 3
oWorksheet.GetRange("D4").SetValue 12.3      ' Установить Price для строки 4
oWorksheet.GetRange("D5").SetValue 24.8      ' Установить Price для строки 5

' Определить диапазон данных для сводной таблицы
Dim dataRef As Object
Set dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5")

' Вставить новую сводную таблицу на новом листе
Dim pivotTable As Object
Set pivotTable = Api.InsertPivotNewWorksheet(dataRef)

' Добавить поля строк в сводную таблицу
pivotTable.AddFields Array("Region", "Style") ' Добавить Region и Style как поля строк

' Добавить поле данных в сводную таблицу
pivotTable.AddDataField "Price" ' Добавить Price как поле данных

' Получить активный лист, где находится сводная таблица
Dim pivotWorksheet As Object
Set pivotWorksheet = Api.GetActiveSheet()

' Получить поле сводной таблицы для Style
Dim pivotField As Object
Set pivotField = pivotTable.GetPivotFields("Style")

' Добавить Region как поле данных в сводной таблице
pivotField.GetTable().AddDataField "Region"
```