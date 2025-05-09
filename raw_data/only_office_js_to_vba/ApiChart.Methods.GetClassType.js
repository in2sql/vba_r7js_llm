**Description / Описание**

This script retrieves the active worksheet, populates specific cells with data, creates a 3D bar chart based on the data, customizes the chart's appearance, and records the chart's class type in a designated cell.

Этот скрипт получает активный лист, заполняет определенные ячейки данными, создает 3D столбчатую диаграмму на основе данных, настраивает внешний вид диаграммы и записывает тип класса диаграммы в указанную ячейку.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells B1, C1, D1
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels in cells A2 and A3
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values for revenue and costs
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the title of the chart
oChart.SetTitle("Financial Overview", 13);

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get the class type of the chart and set it in cell F1
var sClassType = oChart.GetClassType();
oWorksheet.GetRange("F1").SetValue("Class Type: " + sClassType);
```

```vba
' Get the active worksheet
Dim ws As Worksheet
Set ws = ThisWorkbook.ActiveSheet

' Set values in cells B1, C1, D1
ws.Range("B1").Value = 2014
ws.Range("C1").Value = 2015
ws.Range("D1").Value = 2016

' Set labels in cells A2 and A3
ws.Range("A2").Value = "Projected Revenue"
ws.Range("A3").Value = "Estimated Costs"

' Set data values for revenue and costs
ws.Range("B2").Value = 200
ws.Range("B3").Value = 250
ws.Range("C2").Value = 240
ws.Range("C3").Value = 260
ws.Range("D2").Value = 280
ws.Range("D3").Value = 280

' Add a 3D bar chart to the worksheet
Dim cht As Chart
Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 100, 70, 3600, 700).Chart
cht.SetSourceData Source:=ws.Range("A1:D3")
cht.ChartType = xl3DColumn

' Set the title of the chart
cht.HasTitle = True
cht.ChartTitle.Text = "Financial Overview"
cht.ChartTitle.Font.Size = 13

' Set the fill color for the first series
cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)

' Set the fill color for the second series
cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)

' Get the class type of the chart and set it in cell F1
ws.Range("F1").Value = "Class Type: " & TypeName(cht)
```