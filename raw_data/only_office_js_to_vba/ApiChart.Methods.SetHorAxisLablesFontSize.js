# Description

This code sets up a financial overview chart by populating data into cells, adding a 3D bar chart, setting the chart title and axis label font size, and setting the series fill colors.

Этот код настраивает график финансового обзора, заполняя данные в ячейки, добавляя 3D-столбчатую диаграмму, устанавливая заголовок графика и размер шрифта радиальной оси, а также устанавливая цвета заливки серий.

```vba
' Set the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Populate cell values
oWorksheet.Range("B1").Value = 2014
oWorksheet.Range("C1").Value = 2015
oWorksheet.Range("D1").Value = 2016
oWorksheet.Range("A2").Value = "Projected Revenue"
oWorksheet.Range("A3").Value = "Estimated Costs"
oWorksheet.Range("B2").Value = 200
oWorksheet.Range("B3").Value = 250
oWorksheet.Range("C2").Value = 240
oWorksheet.Range("C3").Value = 260
oWorksheet.Range("D2").Value = 280
oWorksheet.Range("D3").Value = 280

' Add a 3D bar chart
Dim oChart As Chart
Set oChart = oWorksheet.Shapes.AddChart2(240, xlBarClustered, 100, 70, 360, 200).Chart

' Set chart title
oChart.HasTitle = True
oChart.ChartTitle.Text = "Financial Overview"
oChart.ChartTitle.Font.Size = 13

' Set horizontal axis label font size
oChart.Axes(xlCategory).TickLabels.Font.Size = 10

' Set series fill colors
oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Series 1
oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Series 2
```

```javascript
// This example sets the font size to the horizontal axis labels.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);
oChart.SetHorAxisLablesFontSize(10);
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false); 
```